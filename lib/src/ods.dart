// Inspired from http://search.cpan.org/~terhechte/Spreadsheet-ReadSXC-0.20/ReadSXC.pm
// table:table format explained here http://books.evc-cit.info/odbook/ch05.html
//
// NOTE: This implementation doesn't support following features
//   - annotations
//   - spanned rows
//   - spanned columns
//   - hidden rows (visible in resulting table)
//   - hidden columns (visible in resulting table)
part of spreadsheet;

/// Read and parse ODS spreadsheet
class OdsDecoder extends SpreadsheetDecoder {
  OdsDecoder(Archive archive) {
    this._archive = archive;
    _tables = new Map<String, SpreadsheetTable>();
    _parseContent();
  }

  _parseContent() {
    var content = _archive.findFile('content.xml');
    content.decompress();
    var document = parse(UTF8.decode(content.content));
    document.findAllElements("table:table").forEach((node) {
      _parseTable(node);
    });
  }

  _parseTable(XmlElement node) {
    var name = node.getAttribute("table:name");
    tables[name] = new SpreadsheetTable();
    ;
    var table = tables[name];

    table._rows = 0;
    node.findElements("table:table-row").forEach((child) {
      _parseRow(child, table);
    });

    // Truncate rows only
    if (table._maxRows == -1) {
      table.rows.clear();
    } else if (table._maxRows < table.rows.length) {
      table.rows.removeRange(table._maxRows, table.rows.length);
    }
    for (var row = 0; row < table.rows.length; row++) {
      if (table._maxCols == -1) {
        table.rows[row].clear();
      } else if (table._maxCols < table.rows[row].length) {
        table.rows[row].removeRange(table._maxCols, table.rows[row].length);
      }
    }
  }

  _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = new List();

    table._cols = 0;
    node.findElements("table:table-cell").forEach((child) {
      _parseCell(child, table, row);
    });

    var repeat = (node.getAttribute("table:number-rows-repeated") != null)
        ? int.parse(node.getAttribute("table:number-rows-repeated"))
        : 1;
    for (var index = 0; index < repeat; index++) {
      table.rows.add(row);
      table._rows++;
    }

    // If row not empty
    if (!row.fold(true, (value, element) => value && (element == null))) {
      if (table._maxRows < table._rows) {
        table._maxRows = table._rows;
      }
    }
  }

  _parseCell(XmlElement node, SpreadsheetTable table, List row) {
    var list = new List<String>();

    node.findElements("text:p").forEach((child) {
      list.add(_parseValue(child));
    });

    var text = (list.isNotEmpty) ? list.join(r'\n').trim() : null;
    var repeat = (node.getAttribute("table:number-columns-repeated") != null)
        ? int.parse(node.getAttribute("table:number-columns-repeated"))
        : 1;
    for (var index = 0; index < repeat; index++) {
      row.add(text);
      table._cols++;
    }

    if (text != null) {
      if (table._maxCols < table._cols) {
        table._maxCols = table._cols;
      }
    }
  }

  _parseValue(XmlElement node) {
    var buffer = new StringBuffer();

    node.children.forEach((child) {
      if (child is XmlElement) {
        buffer.write(_unescape(_parseValue(child)));
      } else if (child is XmlText) {
        buffer.write(_unescape(child.text));
      }
    });

    return buffer.toString();
  }
}
