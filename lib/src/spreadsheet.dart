part of spreadsheet_decoder;

const _spreasheetOds = 'ods';
const _spreasheetXlsx = 'xlsx';
final Map<String, String> _spreasheetExtensionMap = <String, String>{
  _spreasheetOds: 'application/vnd.oasis.opendocument.spreadsheet',
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

String _escape(String text) {
  return text
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&apos;")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;");
}

SpreadsheetDecoder _newSpreadsheetDecoder(Archive archive, bool update) {
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
      return new OdsDecoder(archive, update: update);
    case _spreasheetXlsx:
      return new XlsxDecoder(archive, update: update);
    default:
      throw new UnsupportedError("Spreadsheet format unsupported");
  }
}

// Issue on archive
// https://github.com/brendan-duncan/archive/pull/43
const List<String> _noCompression = const <String>[
  'mimetype',
  'Thumbnails/thumbnail.png',
];

/**
 * Decode a spreadsheet file.
 */
abstract class SpreadsheetDecoder {
  bool _update;
  Archive _archive;
  Map<String, XmlNode> _sheets;
  Map<String, XmlDocument> _xmlFiles;
  Map<String, ArchiveFile> _archiveFiles;

  Map<String, SpreadsheetTable> _tables;

  /// Tables contained in spreadsheet file indexed by their names
  Map<String, SpreadsheetTable> get tables => _tables;

  SpreadsheetDecoder();

  factory SpreadsheetDecoder.decodeBytes(List<int> data, {bool update: false, bool verify: false}) {
    var archive = new ZipDecoder().decodeBytes(data, verify: verify);
    return _newSpreadsheetDecoder(archive, update);
  }

  factory SpreadsheetDecoder.decodeBuffer(InputStream input, {bool update: false, bool verify: false}) {
    var archive = new ZipDecoder().decodeBuffer(input, verify: verify);
    return _newSpreadsheetDecoder(archive, update);
  }

  /// Dump XML content (for debug purpose)
  String dumpXmlContent([String sheet]);

  /// Update the contents from [sheet] of the cell [column]x[row] with ndexes start from 0
  void updateCell(String sheet, int col, int row, dynamic value);

  /// Encode bytes after update
  List<int> encode() {
    if (_update != true) {
      throw new ArgumentError("'update' should be set to 'true' on constructor");
    }

    for (var xmlFile in _xmlFiles.keys) {
      var xml = _xmlFiles[xmlFile].toString();
      var content = UTF8.encode(xml);
      _archiveFiles[xmlFile] = new ArchiveFile(xmlFile, content.length, content);
    }
    return new ZipEncoder().encode(_cloneArchive(_archive));
  }

  Archive _cloneArchive(Archive archive) {
    var clone = new Archive();
    archive.files.forEach((file) {
      if (file.isFile) {
        ArchiveFile copy;
        if (_archiveFiles.containsKey(file.name)) {
          copy = _archiveFiles[file.name];
        } else {
          var content = (file.content as Uint8List).toList();
          //var compress = file.compress;
          var compress = _noCompression.contains(file.name) ? false : true;
          copy = new ArchiveFile(file.name, content.length, content)..compress = compress;
        }
        clone.addFile(copy);
      }
    });
    return clone;
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
