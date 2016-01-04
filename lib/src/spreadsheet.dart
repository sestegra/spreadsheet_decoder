part of spreadsheet_decoder;

const _spreasheetOds  = 'ods';
const _spreasheetXlsx = 'xlsx';
final Map<String, String> _spreasheetExtensionMap = <String, String>{
  _spreasheetOds:  'application/vnd.oasis.opendocument.spreadsheet',
  _spreasheetXlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

final _emptyRow = new List();

// HTML entities to decode (see here http://books.evc-cit.info/apa.php)
// Normalize new line
String _unescape(String text) {
  return text
      .replaceAll("&quot;", '"')
      .replaceAll("&apos;", "'")
      .replaceAll("&amp;", "&")
      .replaceAll("&lt;", "<")
      .replaceAll("&gt;", ">")
      .replaceAll("\r\n", "\n");

}

SpreadsheetDecoder _newSpreadsheetDecoder(Archive archive) {
  // Lookup at file format
  var format;

  // Try OpenDocument format
  var mimetype = archive.findFile('mimetype');
  if (mimetype != null) {
    mimetype.decompress();
    var content = UTF8.decode(mimetype.content);
    if (content == _spreasheetExtensionMap[_spreasheetOds]) {
      format = _spreasheetOds;
    }

  // Try OpenXml Office format
  } else {
    var xl = archive.findFile('xl/workbook.xml');
    format = xl != null ? _spreasheetXlsx : null;
  }

  switch (format) {
    case _spreasheetOds:
      return new OdsDecoder(archive);
    case _spreasheetXlsx:
      return new XlsxDecoder(archive);
    default:
      throw new UnsupportedError("Spreadsheet format unsupported");
  }
}

/**
 * Decode a spreadsheet file.
 */
abstract class SpreadsheetDecoder {
  Archive _archive;
  Map<String, SpreadsheetTable> _tables;
  /// Tables contained in spreadsheet file indexed by their names
  Map<String, SpreadsheetTable> get tables => _tables;

  SpreadsheetDecoder();

  factory SpreadsheetDecoder.decodeBytes(List<int> data, {bool verify: false}) {
    var archive = new ZipDecoder().decodeBytes(data, verify: verify);
    return _newSpreadsheetDecoder(archive);
  }

  factory SpreadsheetDecoder.decodeBuffer(InputStream input, {bool verify: false}) {
    var archive = new ZipDecoder().decodeBuffer(input, verify: verify);
    return _newSpreadsheetDecoder(archive);
  }

  _normalizeTable(SpreadsheetTable table) {
    if (table._maxRows == -1) {
      table._rows.clear();
    } else if (table._maxRows < table._rows.length) {
      table._rows.removeRange(table._maxRows, table._rows.length);
    }
    for (var row = 0; row < table._rows.length; row++) {
      if (table._maxCols == -1) {
        table._rows[row].clear();
      } else if (table._maxCols < table._rows[row].length) {
        table._rows[row].removeRange(table._maxCols, table._rows[row].length);
      } else if (table._maxCols > table._rows[row].length) {
        var repeat = table._maxCols - table._rows[row].length;
        for (var index = 0; index < repeat; index++) {
          table._rows[row].add(null);
        }
      }
    }
  }

  bool _isEmptyRow(List row) {
    return row.fold(true, (value, element) => value && (element == null));
  }

  bool _isNotEmptyRow(List row) {
    return !_isEmptyRow(row);
  }

  _countFilledRow(SpreadsheetTable table, List row) {
    if (_isNotEmptyRow(row)) {
      if (table._maxRows < table._rows.length) {
        table._maxRows = table._rows.length;
      }
    }
  }

  _countFilledColumn(SpreadsheetTable table, List row, dynamic value) {
    if (value != null) {
      if (table._maxCols < row.length) {
        table._maxCols = row.length;
      }
    }
  }
}

/// Table of a spreadsheet file
class SpreadsheetTable {
  int _maxRows = -1;
  int _maxCols = -1;

  List<List> _rows = new List<List>();
  /// List of table's rows
  List<List> get rows => _rows;
}
