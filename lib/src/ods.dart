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
  String get mediaType => "application/vnd.oasis.opendocument.spreadsheet";
  String get extension => ".ods";
  Map<String, List<String>> _styleNames;

  OdsDecoder(Archive archive, {bool update = false}) {
    this._archive = archive;
    this._update = update;
    _tables = Map<String, SpreadsheetTable>();
    _parseContent();
  }

  String dumpXmlContent([String sheet]) {
    if (sheet == null) {
      return _xmlFiles[CONTENT_XML].toXmlString(pretty: true);
    } else {
      return _sheets[sheet].toXmlString(pretty: true);
    }
  }

  void insertColumn(String sheet, int columnIndex) {
    super.insertColumn(sheet, columnIndex);

    var rows = _findRows(_sheets[sheet]).toList();
    for (var row in rows) {
      var cell = _findCellByIndex(row, columnIndex);
      if (cell != null) {
        row.children.insert(row.children.indexOf(cell), _createCell(null));
      } else {
        row.children.add(_createCell(null));
      }
    }
  }

  void removeColumn(String sheet, int columnIndex) {
    super.removeColumn(sheet, columnIndex);

    var rows = _findRows(_sheets[sheet]).toList();
    for (var row in rows) {
      var cell = _findCellByIndex(row, columnIndex);
      row.children.remove(cell);
    }
  }

  void insertRow(String sheet, int rowIndex) {
    super.insertRow(sheet, rowIndex);

    var style = _styleNames['table-row'].first;
    var parent = _sheets[sheet];
    var newRow = _createRow(_tables[sheet]._maxCols, style);
    var row = _findRowByIndex(parent, rowIndex);
    if (row != null) {
      parent.children.insert(parent.children.indexOf(row), newRow);
    } else {
      parent.children.add(newRow);
    }
  }

  void removeRow(String sheet, int rowIndex) {
    super.removeRow(sheet, rowIndex);

    var parent = _sheets[sheet];
    var row = _findRowByIndex(parent, rowIndex);
    parent.children.remove(row);
  }

  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value) {
    super.updateCell(sheet, columnIndex, rowIndex, value);

    var row = _findRowByIndex(_sheets[sheet], rowIndex);
    var cell = _findCellByIndex(row, columnIndex);
    _replaceCell(row, cell, value);
  }

  void _parseContent() {
    var file = _archive.findFile(CONTENT_XML);
    file.decompress();
    var content = parse(utf8.decode(file.content));
    if (_update == true) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlNode>{};
      _xmlFiles = {
        CONTENT_XML: content,
      };
      _parseStyles(content);
    }
    content.findAllElements('table:table').forEach((node) {
      var name = node.getAttribute('table:name');
      if (_update == true) {
        _sheets[name] = node;
      }
      _parseTable(node, name);
    });
  }

  void _parseStyles(XmlDocument document) {
    _styleNames = <String, List<String>>{};
    document.findAllElements('style:style').forEach((style) {
      var name = style.getAttribute('style:name');
      var family = style.getAttribute('style:family');
      _styleNames[family] ??= <String>[];
      _styleNames[family].add(name);
    });
  }

  void _parseTable(XmlElement node, String name) {
    tables[name] = SpreadsheetTable(name);
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
      if (empty) {
        row.parent.children.remove(row);
      }

      return empty;
    });

    filledRows.toList().reversed.forEach((child) {
      _parseRow(child, table);
    });

    _normalizeTable(table);
  }

  void _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = List();
    var cells = _findCells(node);

    // Remove tailing empty cells
    var filledCells = cells.toList().reversed.skipWhile((cell) => _readCell(cell) == null);

    filledCells.toList().reversed.forEach((child) {
      _parseCell(child, table, row);
    });

    var repeat = _getRowRepeated(node);
    for (var index = 0; index < repeat; index++) {
      table._rows.add(List.from(row));
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
        value = value.replaceAll(RegExp('[H|M]'), ':');
        break;
      case 'string':
      default:
        var list = List<String>();
        node.findElements('text:p').forEach((child) {
          list.add(_readString(child));
        });
        value = (list.isNotEmpty) ? list.join('\n') : null;
    }
    return value;
  }

  String _readString(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlElement) {
        buffer.write(_normalizeNewLine(_readString(child)));
      } else if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.text));
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

  static XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
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

    if (row != null) {
      // Expand row if required
      var repeat = _getRowRepeated(row);
      if (repeat != 1) {
        var rows = _expandRepeatedRows(table, row);
        row = rows[rowIndex - (currentIndex - repeat + 1)];
      }
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

    if (cell != null) {
      // Expand cell if required
      var repeat = _getCellRepeated(cell);
      if (repeat != 1) {
        var cells = _expandRepeatedCells(row, cell);
        cell = cells[columnIndex - (currentIndex - repeat + 1)];
      }
    }

    return cell;
  }

  static List<XmlElement> _expandRepeatedRows(XmlElement table, XmlElement row) {
    var repeat = _removeRowRepeated(row);
    var index = table.children.indexOf(row);
    var rows = <XmlElement>[];
    for (var i = 0; i < repeat; i++) {
      rows.add(row.copy());
    }

    table.children
      ..removeAt(index)
      ..insertAll(index, rows);

    return rows;
  }

  static List<XmlElement> _expandRepeatedCells(XmlElement row, XmlElement cell) {
    var repeat = _removeCellRepeated(cell);
    var index = row.children.indexOf(cell);
    var cells = <XmlElement>[];
    for (var i = 0; i < repeat; i++) {
      cells.add(cell.copy());
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

  static XmlElement _createRow(int maxCols, String style) {
    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('table:style-name'), style),
    ];
    var children = <XmlNode>[
      XmlElement(XmlName('table:table-cell'), [
        XmlAttribute(XmlName('table:number-columns-repeated'), maxCols.toString()),
      ]),
    ];
    return XmlElement(XmlName('table:table-row'), attributes, children);
  }

  // TODO Manage value's type
  static XmlElement _createCell(dynamic value) {
    var attributes = value == null
        ? <XmlAttribute>[]
        : <XmlAttribute>[
            XmlAttribute(XmlName('office:value-type'), "string"),
            XmlAttribute(XmlName('calcext:value-type'), "string"),
          ];
    var children = value == null
        ? <XmlNode>[]
        : <XmlNode>[
            XmlElement(XmlName('text:p'), [], [XmlText(value.toString())]),
          ];
    return XmlElement(XmlName('table:table-cell'), attributes, children);
  }
}
