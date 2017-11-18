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

/// Read and parse ODS spreadsheet
class OdsDecoder extends SpreadsheetDecoder {
  OdsDecoder(Archive archive, {bool update = false}) {
    this._archive = archive;
    this._update = update;
    _tables = new Map<String, SpreadsheetTable>();
    _parseContent();
  }

  void updateCell(String sheet, int column, int row, dynamic value) {
    if (_update != true) {
      throw new ArgumentError("'update' should be set to 'true' on constructor");
    }
    if (_sheets.containsKey(sheet) == false) {
      throw new ArgumentError("'$sheet' not found");
    }

    var foundRow = _findRowByIndex(_sheets[sheet], row);
    var foundCell = _findCellByIndex(foundRow, column);
    var foundCellIndex = foundRow.children.indexOf(foundCell);

    // TODO Manage value's type
    var data = _escape(value.toString());
    foundRow.children
      ..removeAt(foundCellIndex)
      ..insert(foundCellIndex, _stringCell(data));

    print(foundRow);
  }

  _parseContent() {
    var file = _archive.findFile('content.xml');
    file.decompress();
    var content = parse(UTF8.decode(file.content));
    if (_update == true) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlNode>{};
      _xmlFiles = {
        'content.xml': content,
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

  _parseTable(XmlElement node, String name) {
    tables[name] = new SpreadsheetTable();
    var table = tables[name];

    _findRows(node).forEach((child) {
      _parseRow(child, table);
    });

    _normalizeTable(table);
  }

  _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = new List();

    _findCells(node).forEach((child) {
      _parseCell(child, table, row);
    });

    var repeat = _getRowRepeated(node);
    for (var index = 0; index < repeat; index++) {
      table._rows.add(row);
    }

    _countFilledRow(table, row);
  }

  _parseCell(XmlElement node, SpreadsheetTable table, List row) {
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
          list.add(_parseString(child));
        });
        value = (list.isNotEmpty) ? list.join('\n') : null;
    }

    var repeat = _getCellRepeated(node);
    for (var index = 0; index < repeat; index++) {
      row.add(value);
    }

    _countFilledColumn(table, row, value);
  }

  _parseString(XmlElement node) {
    var buffer = new StringBuffer();

    node.children.forEach((child) {
      if (child is XmlElement) {
        buffer.write(_unescape(_parseString(child)));
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

  static int _getCellRepeated(XmlElement cell) {
    return (cell.getAttribute('table:number-columns-repeated') != null)
        ? int.parse(cell.getAttribute('table:number-columns-repeated'))
        : 1;
  }

  static XmlElement _findRowByIndex(XmlElement table, int index) {
    XmlElement row;
    var rows = _findRows(table);

    var currentIndex = -1;
    for (var currentRow in rows) {
      currentIndex += _getRowRepeated(currentRow);
      if (currentIndex >= index) {
        row = currentRow;
        break;
      }
    }

    // Expand row if required
    if (currentIndex != index) {
      // TODO Expand repeated rows
    }

    return row;
  }

  static XmlElement _findCellByIndex(XmlElement row, int index) {
    XmlElement cell;
    var cells = _findCells(row);

    var currentIndex = -1;
    for (var currentCell in cells) {
      currentIndex += _getCellRepeated(currentCell);
      if (currentIndex >= index) {
        cell = currentCell;
        break;
      }
    }

    // Expand cell if required
    if (currentIndex != index) {
      // TODO Expand repeated rows
    }

    return cell;
  }

  static XmlElement _stringCell(String value) {
    var attributes = <XmlAttribute>[
      new XmlAttribute(new XmlName('office:value-type'), "string"),
      new XmlAttribute(new XmlName('calcext:value-type'), "string"),
    ];
    var children = <XmlNode>[
      new XmlElement(new XmlName('text:p'), [], [new XmlText(value)]),
    ];
    return new XmlElement(new XmlName('table:table-cell'), attributes, children);
  }
}
