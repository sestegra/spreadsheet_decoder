part of spreadsheet;

const _spreasheetOds  = 'ods';
const _spreasheetXlsx = 'xlsx';
final Map<String, String> _spreasheetExtensionMap = <String, String>{
  _spreasheetOds:  'application/vnd.oasis.opendocument.spreadsheet',
  _spreasheetXlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

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
    var xl = archive.findFile('xl/');
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
 * Decode a spreasheet file.
 */
abstract class SpreadsheetDecoder {
  Archive _archive;
  Map<String, SpreadsheetTable> _tables;
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
}

class SpreadsheetTable {
  int _rows;
  int _cols;
  int _maxRows;
  int _maxCols;

  List<List> _table;
  List<List> get table => _table;
}
