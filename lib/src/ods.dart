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
  OdsDecoder(Archive archive) {
    this._archive = archive;
    _tables = new Map<String, SpreadsheetTable>();
    _parseContent();
  }

  _parseContent() {
    var content = _archive.findFile('content.xml');
    content.decompress();
    var document = parse(UTF8.decode(content.content));
    document.findAllElements('table:table').forEach((node) {
      _parseTable(node);
    });
  }

  _parseTable(XmlElement node) {
    var name = node.getAttribute('table:name');
    tables[name] = new SpreadsheetTable();
    var table = tables[name];

    node.findElements('table:table-row').forEach((child) {
      _parseRow(child, table);
    });

    _normalizeTable(table);
  }

  _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = new List();

    node.findElements('table:table-cell').forEach((child) {
      _parseCell(child, table, row);
    });

    var repeat = (node.getAttribute('table:number-rows-repeated') != null)
        ? int.parse(node.getAttribute('table:number-rows-repeated'))
        : 1;
    for (var index = 0; index < repeat; index++) {
      table._rows.add(row);
    }

    _countFilledRow(table, row);
  }

  _parseCell(XmlElement node, SpreadsheetTable table, List row) {
    var list = new List<String>();

    // Numeric content in table cells
    // http://books.evc-cit.info/odbook/ch05.html
    var value = node.getAttribute('office:value');
    if (value != null) {
      list.add(value);
    } else {
      node.findElements('text:p').forEach((child) {
        list.add(_parseValue(child));
      });
    }

    var text = (list.isNotEmpty) ? list.join('\n').trim() : null;
    var repeat = (node.getAttribute('table:number-columns-repeated') != null)
        ? int.parse(node.getAttribute('table:number-columns-repeated'))
        : 1;
    for (var index = 0; index < repeat; index++) {
      row.add(text);
    }

    _countFilledColumn(table, row, text);
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
