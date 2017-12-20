part of spreadsheet_decoder;

const String _relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const String _relationshipsStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const String _relationshipsWorksheet = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const String _relationshipsSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/// Convert a character based column
int lettersToNumeric(String letters) {
  var sum = 0;
  var mul = 1;
  var n;
  for (var index = letters.length - 1; index >= 0; index--) {
    var c = letters[index].codeUnitAt(0);
    n = 1;
    if (65 <= c && c <= 90) {
      n += c - 65;
    } else if (97 <= c && c <= 122) {
      n += c - 97;
    }
    sum += n * mul;
    mul = mul * 26;
  }
  return sum;
}

/// Convert a number to character based column
String numericToLetters(int number) {
  var letters = '';

  while (number != 0) {
    // Set remainder from 1..26
    var remainder = number % 26;

    if (remainder == 0) {
      remainder = 26;
    }

    // Convert the remainder to a character.
    var letter = new String.fromCharCode(65 + remainder - 1);

    // Accumulate the column letters, right to left.
    letters = letter + letters;

    // Get the next order of magnitude.
    number = (number - 1) ~/ 26;
  }
  return letters;
}

int _letterOnly(int rune) {
  if (65 <= rune && rune <= 90) {
    return rune;
  } else if (97 <= rune && rune <= 122) {
    return rune - 32;
  }
  return 0;
}

// Not used
//int _intOnly(int rune) {
//  if (rune >= 48 && rune < 58) {
//    return rune;
//  }
//  return 0;
//}

String _twoDigits(int n) {
  if (n >= 10) {
    return "${n}";
  }
  return "0${n}";
}

/// Returns the coordinates from a cell name.
/// "A1" returns [1, 1] and the "B3" return [2, 3].
List cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart = UTF8.decode(letters.where((rune) => rune > 0).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);
  var x = lettersToNumeric(lettersPart);
  var y = int.parse(numericsPart);
  return [x, y];
}

/// Read and parse XSLX spreadsheet
class XlsxDecoder extends SpreadsheetDecoder {
  String get mediaType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  String get extension => ".xlsx";

  List<String> _sharedStrings = new List<String>();
  List<int> _numFormats = new List<int>();
  String _stylesTarget;
  String _sharedStringsTarget;
  Map<String, String> _worksheetTargets = new Map<String, String>();

  XlsxDecoder(Archive archive, {bool update = false}) {
    this._archive = archive;
    this._update = update;
    if (_update == true) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlNode>{};
      _xmlFiles = <String, XmlDocument>{};
    }
    _tables = new Map<String, SpreadsheetTable>();
    _parseRelations();
    _parseStyles();
    _parseSharedStrings();
    _parseContent();
  }

  String dumpXmlContent([String sheet]) {
    if (sheet == null) {
      var buffer = new StringBuffer();
      _sheets.forEach((name, document) {
        buffer.writeln(name);
        buffer.writeln(document.toXmlString(pretty: true));
      });
      return buffer.toString();
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

    var foundRow = _findRowByIndex(_sheets[sheet], rowIndex);
    _updateCell(foundRow, columnIndex, rowIndex, value);

    super.updateCell(sheet, columnIndex, rowIndex, value);
  }

  _parseRelations() {
    var relations = _archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = parse(UTF8.decode(relations.content));
      document.findAllElements('Relationship').forEach((node) {
        switch (node.getAttribute('Type')) {
          case _relationshipsStyles:
            _stylesTarget = node.getAttribute('Target');
            break;
          case _relationshipsWorksheet:
            _worksheetTargets[node.getAttribute('Id')] = node.getAttribute('Target');
            break;
          case _relationshipsSharedStrings:
            _sharedStringsTarget = node.getAttribute('Target');
            break;
        }
      });
    }
  }

  _parseStyles() {
    var styles = _archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = parse(UTF8.decode(styles.content));
      document.findAllElements('cellXfs').first.findElements('xf').forEach((node) {
        var numFmtId = node.getAttribute('numFmtId');
        if (numFmtId != null) {
          _numFormats.add(int.parse(numFmtId));
        } else {
          _numFormats.add(0);
        }
      });
    }
  }

  _parseSharedStrings() {
    var sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    if (sharedStrings != null) {
      sharedStrings.decompress();
      var document = parse(UTF8.decode(sharedStrings.content));
      document.findAllElements('si').forEach((node) {
        _parseSharedString(node);
      });
    }
  }

  _parseSharedString(XmlElement node) {
    var list = new List();
    node.findAllElements('t').forEach((child) {
      list.add(_parseValue(child));
    });
    _sharedStrings.add(list.join(''));
  }

  _parseContent() {
    var workbook = _archive.findFile('xl/workbook.xml');
    workbook.decompress();
    var document = parse(UTF8.decode(workbook.content));
    document.findAllElements('sheet').forEach((node) {
      _parseTable(node);
    });
  }

  _parseTable(XmlElement node) {
    var name = node.getAttribute('name');
    var target = _worksheetTargets[node.getAttribute('id', namespace: _relationships)];
    tables[name] = new SpreadsheetTable(name);
    var table = tables[name];

    var file = _archive.findFile("xl/$target");
    file.decompress();

    var content = parse(UTF8.decode(file.content));
    var workbench = content.lastChild as XmlElement;
    var sheet = workbench.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, table);
    });
    if (_update == true) {
      _sheets[name] = sheet;
      _xmlFiles["xl/$target"] = content;
    }

    _normalizeTable(table);
  }

  _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = new List();

    _findCells(node).forEach((child) {
      _parseCell(child, table, row);
    });

    var rowIndex = _getRowNumber(node) - 1;
    if (_isNotEmptyRow(row) && rowIndex > table._rows.length) {
      var repeat = rowIndex - table._rows.length;
      for (var index = 0; index < repeat; index++) {
        table._rows.add(new List());
      }
    }

    if (_isNotEmptyRow(row)) {
      table._rows.add(row);
    } else {
      table._rows.add(new List());
    }

    _countFilledRow(table, row);
  }

  _parseCell(XmlElement node, SpreadsheetTable table, List row) {
    var colIndex = _getCellNumber(node) - 1;
    if (colIndex > row.length) {
      var repeat = colIndex - row.length;
      for (var index = 0; index < repeat; index++) {
        row.add(null);
      }
    }

    if (node.children.isEmpty) {
      return;
    }

    var value;
    var type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        value = _sharedStrings[int.parse(_parseValue(node.findElements('v').first))];
        break;
      // boolean
      case 'b':
        value = _parseValue(node.findElements('v').first) == '1';
        break;
      // formula
      case 'str':
        // <c r="C6" s="1" vm="15" t="str">
        //  <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
        //  <v>2838512.355</v>
        // </c>
        value = _parseValue(node.findElements('v').first);
        break;
      // inline string
      case 'inlineStr':
        // <c r="B2" t="inlineStr">
        // <is><t>Hello world</t></is>
        // </c>
        value = _parseValue(node.findAllElements('t').first);
        break;
      // number
      case 'n':
      default:
        var s = node.getAttribute('s');
        var valueNode = node.findElements('v');
        var content = valueNode.first;
        if (s != null) {
          var fmtId = _numFormats[int.parse(s)];
          // date
          if (((fmtId >= 14) && (fmtId <= 17)) || (fmtId == 22)) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = new DateTime(1899, 12, 30);
            value = date.add(new Duration(milliseconds: delta.toInt())).toIso8601String();
            // time
          } else if (((fmtId >= 18) && (fmtId <= 21)) || ((fmtId >= 45) && (fmtId <= 47))) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = new DateTime(0);
            date = date.add(new Duration(milliseconds: delta.toInt()));
            value = "${_twoDigits(date.hour)}:${_twoDigits(date.minute)}:${_twoDigits(date.second)}";
            // number
          } else {
            value = num.parse(_parseValue(content));
          }
        } else {
          value = num.parse(_parseValue(content));
        }
    }
    row.add(value);

    _countFilledColumn(table, row, value);
  }

  _parseValue(XmlElement node) {
    var buffer = new StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) {
        buffer.write(_unescape(child.text));
      }
    });

    return buffer.toString();
  }

  static Iterable<XmlElement> _findRows(XmlElement table) => table.findElements('row');

  static Iterable<XmlElement> _findCells(XmlElement row) => row.findElements('c');

  static int _getRowNumber(XmlElement row) => int.parse(row.getAttribute('r'));

  static int _getCellNumber(XmlElement cell) {
    var coords = cellCoordsFromCellId(cell.getAttribute('r'));
    return coords[0];
  }

  static XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
    XmlElement row;
    var rows = _findRows(table);

    var currentIndex = 0;
    for (var currentRow in rows) {
      currentIndex = _getRowNumber(currentRow) - 1;
      if (currentIndex >= rowIndex) {
        row = currentRow;
        break;
      }
    }

    // Create row if required
    if (row == null || currentIndex != rowIndex) {
      row = _insertRow(table, row, rowIndex);
    }

    return row;
  }

  static XmlElement _updateCell(XmlElement node, int columnIndex, int rowIndex, dynamic value) {
    XmlElement cell;
    var cells = _findCells(node);

    var currentIndex = 0; // cells could be empty
    for (var currentCell in cells) {
      currentIndex = _getCellNumber(currentCell) - 1;
      if (currentIndex >= columnIndex) {
        cell = currentCell;
        break;
      }
    }

    if (cell == null || currentIndex != columnIndex) {
      cell = _insertCell(node, cell, columnIndex, rowIndex, value);
    } else {
      cell = _replaceCell(node, cell, columnIndex, rowIndex, value);
    }

    return cell;
  }

  static XmlElement _createRow(int rowIndex) {
    var attributes = <XmlAttribute>[
      new XmlAttribute(new XmlName('r'), (rowIndex + 1).toString()),
    ];
    return new XmlElement(new XmlName('row'), attributes, []);
  }

  static XmlElement _insertRow(XmlElement table, XmlElement lastRow, int rowIndex) {
    var row = _createRow(rowIndex);
    if (lastRow == null) {
      table.children.add(row);
    } else {
      var index = table.children.indexOf(lastRow);
      table.children.insert(index, row);
    }
    return row;
  }

  static XmlElement _insertCell(XmlElement row, XmlElement lastCell, int columnIndex, int rowIndex, dynamic value) {
    var cell = _createCell(columnIndex, rowIndex, value);
    if (lastCell == null) {
      row.children.add(cell);
    } else {
      var index = row.children.indexOf(lastCell);
      row.children.insert(index, cell);
    }
    return cell;
  }

  static XmlElement _replaceCell(XmlElement row, XmlElement lastCell, int columnIndex, int rowIndex, dynamic value) {
    var index = lastCell == null ? 0 : row.children.indexOf(lastCell);
    var cell = _createCell(columnIndex, rowIndex, value);
    row.children
      ..removeAt(index)
      ..insert(index, cell);
    return cell;
  }

  // TODO Manage value's type
  static XmlElement _createCell(int columnIndex, int rowIndex, dynamic value) {
    var attributes = <XmlAttribute>[
      new XmlAttribute(new XmlName('r'), '${numericToLetters(columnIndex + 1)}${rowIndex + 1}'),
      new XmlAttribute(new XmlName('t'), 'inlineStr'),
    ];
    var children = <XmlElement>[
      new XmlElement(new XmlName('is'), [], [
        new XmlElement(new XmlName('t'), [], [new XmlText(_escape(value.toString()))])
      ]),
    ];
    return new XmlElement(new XmlName('c'), attributes, children);
  }
}
