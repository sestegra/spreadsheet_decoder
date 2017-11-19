// Inspired from http://search.cpan.org/~terhechte/Spreadsheet-ReadSXC-0.20/ReadSXC.pm
// table:table format explained here http://books.evc-cit.info/odbook/ch05.html
//
// NOTE: This implementation doesn't support following features
//   - annotations
//   - spanned rows
//   - spanned columns
//   - hidden rows (visible in resulting table)
//   - hidden columns (visible in resulting table)
part of spreadsheet_decoder;

const String CONTENT_XML = 'content.xml';

/// Read and parse ODS spreadsheet
class OdsDecoder extends SpreadsheetDecoder {
  OdsDecoder(Archive archive, {bool update = false}) {
    this._archive = archive;
    this._update = update;
    _tables = new Map<String, SpreadsheetTable>();
    _parseContent();
  }

  String dumpXmlContent([String sheet]) {
    if (sheet == null) {
      return _xmlFiles[CONTENT_XML].toXmlString(pretty: true);
    } else {
      return _sheets[sheet].toXmlString(pretty: true);
    }
  }

  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value) {
    if (_update != true) {
      throw new ArgumentError("'update' should be set to 'true' on constructor");
    }
    if (_sheets.containsKey(sheet) == false) {
      throw new ArgumentError("'$sheet' not found");
    }

    var row = _findRowByIndex(_sheets[sheet], rowIndex);
    var cell = _findCellByIndex(row, columnIndex);
    _replaceCell(row, cell, value);

    super.updateCell(sheet, columnIndex, rowIndex, value);
  }

  void _parseContent() {
    var file = _archive.findFile(CONTENT_XML);
    file.decompress();
    var content = _parseXml(UTF8.decode(file.content));
    if (_update == true) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlNode>{};
      _xmlFiles = {
        CONTENT_XML: content,
      };
    }
    content.findAllElements('table:table').forEach((node) {
      var name = node.getAttribute('table:name');
      if (_update == true) {
        _sheets[name] = node;
      }
      _parseTable(node, name);
    });
  }

  void _parseTable(XmlElement node, String name) {
    tables[name] = new SpreadsheetTable();
    var table = tables[name];
    var rows = _findRows(node);

    // Remove tailing empty rows
    var filledRows = rows.toList().reversed.skipWhile((row) {
      var empty = true;
      for (var cell in _findCells(row)) {
        if (_readCell(cell) != null) {
          empty = false;
          break;
        }
      }
      return empty;
    });

    filledRows.toList().reversed.forEach((child) {
      _parseRow(child, table);
    });

    _normalizeTable(table);
  }

  void _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = new List();
    var cells = _findCells(node);

    // Remove tailing empty cells
    var filledCells = cells.toList().reversed.skipWhile((cell) => _readCell(cell) == null)  ;

    filledCells.toList().reversed.forEach((child) {
      _parseCell(child, table, row);
    });

    var repeat = _getRowRepeated(node);
    for (var index = 0; index < repeat; index++) {
      table._rows.add(new List.from(row));
    }

    _countFilledRow(table, row);
  }

  void _parseCell(XmlElement node, SpreadsheetTable table, List row) {
    var value = _readCell(node);
    var repeat = _getCellRepeated(node);
    for (var index = 0; index < repeat; index++) {
      row.add(value);
    }

    _countFilledColumn(table, row, value);
  }

  dynamic _readCell(XmlElement node) {
    var value;
    var type = node.getAttribute('office:value-type');
    switch (type) {
      case 'float':
      case 'percentage':
      case 'currency':
        value = num.parse(node.getAttribute('office:value'));
        break;
      case 'boolean':
        value = node.getAttribute('office:boolean-value').toLowerCase() == 'true';
        break;
      case 'date':
        value = DateTime.parse(node.getAttribute('office:date-value')).toIso8601String();
        break;
      case 'time':
        value = node.getAttribute('office:time-value');
        value = value.substring(2, value.length - 1);
        value = value.replaceAll(new RegExp('[H|M]'), ':');
        break;
      case 'string':
      default:
        var list = new List<String>();
        node.findElements('text:p').forEach((child) {
          list.add(_readString(child));
        });
        value = (list.isNotEmpty) ? list.join('\n') : null;
    }
    return value;
  }

  String _readString(XmlElement node) {
    var buffer = new StringBuffer();

    node.children.forEach((child) {
      if (child is XmlElement) {
        buffer.write(_unescape(_readString(child)));
      } else if (child is XmlText) {
        buffer.write(_unescape(child.text));
      }
    });

    return buffer.toString();
  }

  static Iterable<XmlElement> _findRows(XmlElement table) => table.findElements('table:table-row');

  static Iterable<XmlElement> _findCells(XmlElement row) => row.findElements('table:table-cell');

  static int _getRowRepeated(XmlElement row) {
    return (row.getAttribute('table:number-rows-repeated') != null)
        ? int.parse(row.getAttribute('table:number-rows-repeated'))
        : 1;
  }

  static int _removeRowRepeated(XmlElement row) {
    var node = row.getAttributeNode('table:number-rows-repeated');
    row.attributes.remove(node);
    return int.parse(node.value);
  }

  static int _getCellRepeated(XmlElement cell) {
    return (cell.getAttribute('table:number-columns-repeated') != null)
        ? int.parse(cell.getAttribute('table:number-columns-repeated'))
        : 1;
  }

  static int _removeCellRepeated(XmlElement cell) {
    var node = cell.getAttributeNode('table:number-columns-repeated');
    cell.attributes.remove(node);
    return int.parse(node.value);
  }

  XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
    XmlElement row;
    var rows = _findRows(table);

    var currentIndex = -1;
    for (var currentRow in rows) {
      currentIndex += _getRowRepeated(currentRow);
      if (currentIndex >= rowIndex) {
        row = currentRow;
        break;
      }
    }

    // Expand row if required
    var repeat = _getRowRepeated(row);
    if (repeat != 1) {
      var rows = _expandRepeatedRows(table, row);
      row = rows[rowIndex - (currentIndex - repeat + 1)];
    }

    return row;
  }

  static XmlElement _findCellByIndex(XmlElement row, int columnIndex) {
    XmlElement cell;
    var cells = _findCells(row);

    var currentIndex = -1;
    for (var currentCell in cells) {
      currentIndex += _getCellRepeated(currentCell);
      if (currentIndex >= columnIndex) {
        cell = currentCell;
        break;
      }
    }

    // Expand cell if required
    var repeat = _getCellRepeated(cell);
    if (repeat != 1) {
      var cells = _expandRepeatedCells(row, cell);
      cell = cells[columnIndex - (currentIndex - repeat + 1)];
    }

    return cell;
  }

  List<XmlElement> _expandRepeatedRows(XmlElement table, XmlElement row) {
    var repeat = _removeRowRepeated(row);
    var index = table.children.indexOf(row);
    var xml = row.toString();
    var rows = <XmlElement>[];
    for (var i = 0; i < repeat; i++) {
      rows.add(_parseXml(xml).document.root.children.first);
    }

    table.children
      ..removeAt(index)
      ..insertAll(index, rows);

    return rows;
  }

  static List<XmlElement> _expandRepeatedCells(XmlElement row, XmlElement cell) {
    var repeat = _removeCellRepeated(cell);
    var index = row.children.indexOf(cell);
    var xml = cell.toString();
    var cells = <XmlElement>[];
    for (var i = 0; i < repeat; i++) {
      cells.add(parse(xml).document.root.children.first);
    }

    row.children
      ..removeAt(index)
      ..insertAll(index, cells);

    return cells;
  }

  static XmlElement _replaceCell(XmlElement row, XmlElement lastCell, dynamic value) {
    var index = row.children.indexOf(lastCell);
    var cell = _createCell(value);
    row.children
      ..removeAt(index)
      ..insert(index, cell);
    return cell;
  }

  // TODO Manage value's type
  static XmlElement _createCell(String value) {
    var attributes = <XmlAttribute>[
      new XmlAttribute(new XmlName('office:value-type'), "string"),
      new XmlAttribute(new XmlName('calcext:value-type'), "string"),
    ];
    var children = <XmlNode>[
      new XmlElement(new XmlName('text:p'), [], [new XmlText(_escape(value.toString()))]),
    ];
    return new XmlElement(new XmlName('table:table-cell'), attributes, children);
  }
}
