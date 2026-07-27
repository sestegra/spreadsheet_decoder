part of spreadsheet_decoder;

const String _relationships =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const String _relationshipsStyles =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const String _relationshipsWorksheet =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const String _relationshipsSharedStrings =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';

/// Convert a character based column
int lettersToNumeric(String letters) {
  var sum = 0;
  var mul = 1;
  for (var index = letters.length - 1; index >= 0; index--) {
    var c = letters[index].codeUnitAt(0);
    var n = 1;
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
    var letter = String.fromCharCode(65 + remainder - 1);

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
    return '$n';
  }
  return '0$n';
}

String _fourDigits(int n) {
  var v = n.abs().toString().padLeft(4, '0');
  return n < 0 ? '-$v' : v;
}

const List<String> _monthNames = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December' //
];

const List<String> _dayNames = [
  'Monday', 'Tuesday', 'Wednesday', 'Thursday',
  'Friday', 'Saturday', 'Sunday' //
];

/// Strip literal/escaped portions of an Excel number format code so token
/// scanning only sees real format tokens. Removes "..." literals,
/// [..] bracket sections (colors, conditions, locale modifiers like
/// `[$-409]`) and `\X` escape pairs.
String _stripFormatLiterals(String code) {
  var buf = StringBuffer();
  var i = 0;
  while (i < code.length) {
    var ch = code[i];
    if (ch == '"') {
      i++;
      while (i < code.length && code[i] != '"') {
        i++;
      }
      if (i < code.length) i++;
      continue;
    }
    if (ch == r'\' && i + 1 < code.length) {
      i += 2;
      continue;
    }
    if (ch == '[') {
      while (i < code.length && code[i] != ']') {
        i++;
      }
      if (i < code.length) i++;
      continue;
    }
    buf.write(ch);
    i++;
  }
  return buf.toString();
}

/// Returns true if the given Excel format code represents a date or
/// date+time value (contains year or day tokens).
bool _isDateTimeFormatCode(String code) {
  var stripped = _stripFormatLiterals(code).toLowerCase();
  var section = stripped.split(';').first;
  return section.contains('y') || section.contains('d');
}

/// Returns true if [fmtId] is a built-in (implicit) number format that Excel
/// renders as a date, meaning the cell value holds a date serial.
///
/// Besides the well-known western formats (14-17, 22) this covers the
/// locale-specific ranges the spec reserves for CJK (27-36, 50-58) and Thai
/// (71-81) date formats. Those ids carry no `<numFmt>` entry in the workbook,
/// so they have to be recognised by id alone.
bool _isBuiltInDateFormat(int fmtId) {
  if (((fmtId >= 14) && (fmtId <= 17)) || (fmtId == 22)) return true;
  // CJK: yyyy年m月, m月d日, yyyy年m月d日, era-based ge.m.d, ...
  if ((fmtId >= 27) && (fmtId <= 31)) return true;
  if (fmtId == 36) return true;
  if ((fmtId >= 50) && (fmtId <= 54)) return true;
  if ((fmtId == 57) || (fmtId == 58)) return true;
  // Thai: ว/ด/ปปปป, ว-ดดด-ปป, d/m/bb, ...
  if ((fmtId >= 71) && (fmtId <= 74)) return true;
  if ((fmtId == 77) || (fmtId == 81)) return true;
  return false;
}

/// Returns true if [fmtId] is a built-in number format that Excel renders as a
/// time of day only (no date part). Companion to [_isBuiltInDateFormat].
bool _isBuiltInTimeFormat(int fmtId) {
  if ((fmtId >= 18) && (fmtId <= 21)) return true;
  // CJK: h時mm分, 上午/下午h時mm分ss秒, ...
  if ((fmtId >= 32) && (fmtId <= 35)) return true;
  if ((fmtId >= 45) && (fmtId <= 47)) return true;
  if ((fmtId == 55) || (fmtId == 56)) return true;
  // Thai: ช:นน, นน:ทท.0, ...
  if ((fmtId == 75) || (fmtId == 76)) return true;
  if ((fmtId >= 78) && (fmtId <= 80)) return true;
  return false;
}

/// Returns true if the format code represents a time-only value
/// (hours/minutes/seconds without a date part).
bool _isTimeOnlyFormatCode(String code) {
  var stripped = _stripFormatLiterals(code).toLowerCase();
  var section = stripped.split(';').first;
  if (section.contains('y') || section.contains('d')) return false;
  return section.contains('h') || section.contains('s');
}

/// Format a [DateTime] using a (subset of) Excel number format tokens.
/// Supports y/yy/yyyy, m/mm/mmm/mmmm/mmmmm (month or minute by context),
/// d/dd/ddd/dddd, h/hh, s/ss, AM/PM and A/P. Literal characters and
/// quoted text pass through unchanged.
String _formatDateTimeWithCode(DateTime date, String code) {
  var section = code.split(';').first;
  var lower = section.toLowerCase();
  var twelveHour = lower.contains('am/pm') || lower.contains('a/p');
  var hour12 = ((date.hour + 11) % 12) + 1;

  var buf = StringBuffer();
  var i = 0;
  var lastWasHour = false;
  while (i < section.length) {
    var ch = section[i];
    var lc = ch.toLowerCase();

    if (ch == '"') {
      i++;
      while (i < section.length && section[i] != '"') {
        buf.write(section[i]);
        i++;
      }
      if (i < section.length) i++;
      continue;
    }
    if (ch == r'\' && i + 1 < section.length) {
      buf.write(section[i + 1]);
      i += 2;
      continue;
    }
    if (ch == '[') {
      var end = section.indexOf(']', i + 1);
      if (end == -1) {
        i++;
        continue;
      }
      var inner = section.substring(i + 1, end).toLowerCase();
      if (inner == 'h' || inner == 'hh') {
        buf.write(date.hour.toString());
      } else if (inner == 'm' || inner == 'mm') {
        buf.write(date.minute.toString());
      } else if (inner == 's' || inner == 'ss') {
        buf.write(date.second.toString());
      }
      i = end + 1;
      continue;
    }

    // AM/PM marker (check before single-char fallthrough)
    if (i + 5 <= section.length &&
        section.substring(i, i + 5).toLowerCase() == 'am/pm') {
      var src = section.substring(i, i + 5);
      var meridiem = date.hour < 12 ? 'AM' : 'PM';
      if (src == src.toLowerCase()) meridiem = meridiem.toLowerCase();
      buf.write(meridiem);
      i += 5;
      continue;
    }
    if (i + 3 <= section.length &&
        section.substring(i, i + 3).toLowerCase() == 'a/p') {
      var src = section.substring(i, i + 3);
      var meridiem = date.hour < 12 ? 'A' : 'P';
      if (src == src.toLowerCase()) meridiem = meridiem.toLowerCase();
      buf.write(meridiem);
      i += 3;
      continue;
    }

    if ('ymdhs'.contains(lc)) {
      var run = lc;
      var j = i + 1;
      while (j < section.length && section[j].toLowerCase() == lc) {
        run += lc;
        j++;
      }
      switch (lc) {
        case 'y':
          if (run.length >= 4) {
            buf.write(_fourDigits(date.year));
          } else {
            buf.write(_twoDigits(date.year % 100));
          }
          lastWasHour = false;
          break;
        case 'd':
          if (run.length == 1) {
            buf.write(date.day.toString());
          } else if (run.length == 2) {
            buf.write(_twoDigits(date.day));
          } else if (run.length == 3) {
            buf.write(_dayNames[(date.weekday - 1) % 7].substring(0, 3));
          } else {
            buf.write(_dayNames[(date.weekday - 1) % 7]);
          }
          lastWasHour = false;
          break;
        case 'h':
          var h = twelveHour ? hour12 : date.hour;
          if (run.length == 1) {
            buf.write(h.toString());
          } else {
            buf.write(_twoDigits(h));
          }
          lastWasHour = true;
          break;
        case 's':
          if (run.length == 1) {
            buf.write(date.second.toString());
          } else {
            buf.write(_twoDigits(date.second));
          }
          lastWasHour = false;
          break;
        case 'm':
          var isMinute = lastWasHour;
          if (!isMinute) {
            var k = j;
            while (k < section.length && section[k] == ' ') {
              k++;
            }
            if (k < section.length && section[k].toLowerCase() == 's') {
              isMinute = true;
            }
          }
          if (isMinute) {
            if (run.length == 1) {
              buf.write(date.minute.toString());
            } else {
              buf.write(_twoDigits(date.minute));
            }
          } else {
            if (run.length == 1) {
              buf.write(date.month.toString());
            } else if (run.length == 2) {
              buf.write(_twoDigits(date.month));
            } else if (run.length == 3) {
              buf.write(_monthNames[date.month - 1].substring(0, 3));
            } else if (run.length == 5) {
              buf.write(_monthNames[date.month - 1].substring(0, 1));
            } else {
              buf.write(_monthNames[date.month - 1]);
            }
          }
          lastWasHour = false;
          break;
      }
      i = j;
      continue;
    }

    buf.write(ch);
    i++;
  }
  return buf.toString();
}

/// Returns the coordinates from a cell name.
/// "A1" returns [1, 1] and the "B3" return [2, 3].
List cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart =
      utf8.decode(letters.where((rune) => rune > 0).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);
  var x = lettersToNumeric(lettersPart);
  var y = int.parse(numericsPart);
  return [x, y];
}

/// Read and parse XSLX spreadsheet
class XlsxDecoder extends SpreadsheetDecoder {
  @override
  String get mediaType =>
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  @override
  String get extension => '.xlsx';

  final List<String> _sharedStrings = <String>[];
  final List<int> _numFormats = <int>[];
  // Custom number format codes from <numFmts> (numFmtId >= 164 typically,
  // but spec allows any id when overridden). Maps numFmtId -> formatCode.
  final Map<int, String> _customNumFormats = <int, String>{};
  String? _stylesTarget;
  String? _sharedStringsTarget;
  final Map<String, String> _worksheetTargets = <String, String>{};

  XlsxDecoder(Archive archive,
      {bool update = false,
      String dateFormat = SpreadsheetDecoder.defaultDateFormat}) {
    _archive = archive;
    _update = update;
    _dateFormat = dateFormat;
    if (_update == true) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlElement>{};
      _xmlFiles = <String, XmlDocument>{};
    }
    _tables = <String, SpreadsheetTable>{};
    _parseRelations();
    _parseStyles();
    _parseSharedStrings();
    _parseContent();
  }

  @override
  String dumpXmlContent([String? sheet]) {
    if (sheet == null) {
      var buffer = StringBuffer();
      _sheets.forEach((name, document) {
        buffer.writeln(name);
        buffer.writeln(document.toXmlString(pretty: true));
      });
      return buffer.toString();
    } else {
      return _sheets[sheet]!.toXmlString(pretty: true);
    }
  }

  @override
  void insertColumn(String sheet, int columnIndex) {
    super.insertColumn(sheet, columnIndex);

    for (var row in _findRows(_sheets[sheet]!)) {
      XmlElement? cell;
      var cells = _findCells(row);

      var currentIndex = 0; // cells could be empty
      for (var currentCell in cells) {
        currentIndex = _getCellNumber(currentCell) - 1;
        if (currentIndex >= columnIndex) {
          cell = currentCell;
          break;
        }
      }

      if (cell != null) {
        cells
            .skipWhile((c) => c != cell)
            .forEach((c) => _setCellColNumber(c, _getCellNumber(c) + 1));
      }
      // Nothing to do if cell == null
    }
  }

  @override
  void removeColumn(String sheet, int columnIndex) {
    super.removeColumn(sheet, columnIndex);

    for (var row in _findRows(_sheets[sheet]!)) {
      XmlElement? cell;
      var cells = _findCells(row);

      var currentIndex = 0; // cells could be empty
      for (var currentCell in cells) {
        currentIndex = _getCellNumber(currentCell) - 1;
        if (currentIndex >= columnIndex) {
          cell = currentCell;
          break;
        }
      }

      if (cell != null) {
        cells
            .skipWhile((c) => c != cell)
            .forEach((c) => _setCellColNumber(c, _getCellNumber(c) - 1));
        cell.parent!.children.remove(cell);
      }
    }
  }

  @override
  void insertRow(String sheet, int rowIndex) {
    super.insertRow(sheet, rowIndex);

    var parent = _sheets[sheet]!;
    if (rowIndex < _tables[sheet]!._maxRows - 1) {
      var foundRow = _findRowByIndex(_sheets[sheet]!, rowIndex);
      _insertRow(parent, foundRow, rowIndex);
      parent.children
          .whereType<XmlElement>()
          .skipWhile((row) => row != foundRow)
          .forEach((row) {
        var rIndex = _getRowNumber(row) + 1;
        _setRowNumber(row, rIndex);
        _findCells(row).forEach((cell) {
          _setCellRowNumber(cell, rIndex);
        });
      });
    } else {
      _insertRow(parent, null, rowIndex);
    }
  }

  @override
  void removeRow(String sheet, int rowIndex) {
    super.removeRow(sheet, rowIndex);

    var parent = _sheets[sheet]!;
    var foundRow = _findRowByIndex(parent, rowIndex);
    parent.children
        .whereType<XmlElement>()
        .skipWhile((row) => row != foundRow)
        .forEach((row) {
      var rIndex = _getRowNumber(row) - 1;
      _setRowNumber(row, rIndex);
      _findCells(row).forEach((cell) {
        _setCellRowNumber(cell, rIndex);
      });
    });
    parent.children.remove(foundRow);
  }

  @override
  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value) {
    super.updateCell(sheet, columnIndex, rowIndex, value);

    var foundRow = _findRowByIndex(_sheets[sheet]!, rowIndex);
    _updateCell(foundRow, columnIndex, rowIndex, value);
  }

  void _parseRelations() {
    var relations = _archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = XmlDocument.parse(utf8.decode(relations.content));
      document.findAllElements('Relationship').forEach((node) {
        var attr = node.getAttribute('Target');
        switch (node.getAttribute('Type')) {
          case _relationshipsStyles:
            _stylesTarget = attr!;
            break;
          case _relationshipsWorksheet:
            _worksheetTargets[node.getAttribute('Id')!] = attr!;
            break;
          case _relationshipsSharedStrings:
            _sharedStringsTarget = attr!;
            break;
        }
      });
    }
  }

  void _parseStyles() {
    var styles = _archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = XmlDocument.parse(utf8.decode(styles.content));

      // Parse custom numFmts (formatCode strings keyed by numFmtId).
      var numFmtsElems = document.findAllElements('numFmts');
      if (numFmtsElems.isNotEmpty) {
        numFmtsElems.first.findElements('numFmt').forEach((node) {
          var idAttr = node.getAttribute('numFmtId');
          var codeAttr = node.getAttribute('formatCode');
          if (idAttr != null && codeAttr != null) {
            _customNumFormats[int.parse(idAttr)] = codeAttr;
          }
        });
      }

      document
          .findAllElements('cellXfs')
          .first
          .findElements('xf')
          .forEach((node) {
        var numFmtId = node.getAttribute('numFmtId');
        if (numFmtId != null) {
          _numFormats.add(int.parse(numFmtId));
        } else {
          _numFormats.add(0);
        }
      });
    }
  }

  void _parseSharedStrings() {
    var sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    if (sharedStrings != null) {
      sharedStrings.decompress();
      var document = XmlDocument.parse(utf8.decode(sharedStrings.content));
      document.findAllElements('si').forEach((node) {
        _parseSharedString(node);
      });
    }
  }

  String _parseRichText(XmlElement node) {
    return _parseValue(node.findElements('t').first);
  }

  void _parseSharedString(XmlElement node) {
    var list = [];
    node.childElements.forEach((node) {
      if (node.localName == 't') {
        list.add(_parseValue(node));
      } else if (node.localName == 'r') {
        list.add(_parseRichText(node));
      } else {
        // ignores <rPh> and <phoneticPr>
      }
    });
    _sharedStrings.add(list.join(''));
  }

  void _parseContent() {
    var workbook = _archive.findFile('xl/workbook.xml');
    if (workbook == null) {
      throw FormatException('Missing required file: xl/workbook.xml');
    }
    workbook.decompress();
    var document = XmlDocument.parse(utf8.decode(workbook.content));
    document.findAllElements('sheet').forEach((node) {
      _parseTable(node);
    });
  }

  void _parseTable(XmlElement node) {
    var name = node.getAttribute('name')!;
    var target =
        _worksheetTargets[node.getAttribute('id', namespace: _relationships)]!;
    var table = tables[name] = SpreadsheetTable(name);

    final namePath =
        target.startsWith('/') ? target.substring(1) : 'xl/$target';
    var file = _archive.findFile(namePath);
    if (file == null) {
      throw FormatException('Missing required file: $namePath');
    }
    file.decompress();

    var content = XmlDocument.parse(utf8.decode(file.content));
    var worksheet = content.findElements('worksheet').first;
    var sheet = worksheet.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, table);
    });
    if (_update == true) {
      _sheets[name] = sheet;
      _xmlFiles[namePath] = content;
    }

    _normalizeTable(table);
  }

  void _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = [];

    _findCells(node).forEach((child) {
      _parseCell(child, table, row);
    });

    var rowIndex = _getRowNumber(node) - 1;
    if (_isNotEmptyRow(row) && rowIndex > table._rows.length) {
      var repeat = rowIndex - table._rows.length;
      for (var index = 0; index < repeat; index++) {
        table._rows.add([]);
      }
    }

    if (_isNotEmptyRow(row)) {
      table._rows.add(row);
    } else {
      table._rows.add([]);
    }

    _countFilledRow(table, row);
  }

  void _parseCell(XmlElement node, SpreadsheetTable table, List row) {
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

    dynamic value;
    var type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        value = _sharedStrings[
            int.parse(_parseValue(node.findElements('v').first))];
        break;
      // boolean
      case 'b':
        value = _parseValue(node.findElements('v').first) == '1';
        break;
      // error
      case 'e':
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
          if (_isBuiltInDateFormat(fmtId)) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(1899, 12, 30)
                .add(Duration(milliseconds: delta.toInt()));
            value = _formatDateTimeWithCode(date, _dateFormat);
            // time
          } else if (_isBuiltInTimeFormat(fmtId)) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(0);
            date = date.add(Duration(milliseconds: delta.toInt()));
            value =
                '${_twoDigits(date.hour)}:${_twoDigits(date.minute)}:${_twoDigits(date.second)}';
            // number
          } else if (_customNumFormats.containsKey(fmtId)) {
            // Custom number format declared in <numFmts>. If the caller
            // supplied a non-default [dateFormat], honour it for date
            // cells; otherwise fall back to the workbook's own format
            // code so the visible output matches Excel.
            var workbookCode = _customNumFormats[fmtId]!;
            var serial = num.parse(_parseValue(content));
            if (_isDateTimeFormatCode(workbookCode)) {
              var delta = serial * 24 * 3600 * 1000;
              var date = DateTime(1899, 12, 30)
                  .add(Duration(milliseconds: delta.toInt()));
              var code = _dateFormat == SpreadsheetDecoder.defaultDateFormat
                  ? workbookCode
                  : _dateFormat;
              value = _formatDateTimeWithCode(date, code);
            } else if (_isTimeOnlyFormatCode(workbookCode)) {
              var delta = serial * 24 * 3600 * 1000;
              var date = DateTime(0).add(Duration(milliseconds: delta.toInt()));
              value = _formatDateTimeWithCode(date, workbookCode);
            } else {
              value = serial;
            }
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

  String _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    for (var child in node.children) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.text));
      }
    }

    return buffer.toString();
  }

  static Iterable<XmlElement> _findRows(XmlElement table) =>
      table.findElements('row');

  static Iterable<XmlElement> _findCells(XmlElement row) =>
      row.findElements('c');

  static int _getRowNumber(XmlElement row) => int.parse(row.getAttribute('r')!);
  static void _setRowNumber(XmlElement row, int index) =>
      row.getAttributeNode('r')!.value = index.toString();

  static int _getCellNumber(XmlElement cell) {
    var coords = cellCoordsFromCellId(cell.getAttribute('r')!);
    return coords[0];
  }

  static void _setCellColNumber(XmlElement cell, int colIndex) {
    var attr = cell.getAttributeNode('r')!;
    var coords = cellCoordsFromCellId(attr.value);
    attr.value = '${numericToLetters(colIndex)}${coords[1]}';
  }

  static void _setCellRowNumber(XmlElement cell, int rowIndex) {
    var attr = cell.getAttributeNode('r')!;
    var coords = cellCoordsFromCellId(attr.value);
    attr.value = '${numericToLetters(coords[0])}$rowIndex';
  }

  static XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
    XmlElement? row;
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

  static XmlElement _updateCell(
      XmlElement node, int columnIndex, int rowIndex, dynamic value) {
    XmlElement? cell;
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
      XmlAttribute(XmlName('r'), (rowIndex + 1).toString()),
    ];
    return XmlElement(XmlName('row'), attributes, []);
  }

  static XmlElement _insertRow(
      XmlElement table, XmlElement? lastRow, int rowIndex) {
    var row = _createRow(rowIndex);
    if (lastRow == null) {
      table.children.add(row);
    } else {
      var index = table.children.indexOf(lastRow);
      table.children.insert(index, row);
    }
    return row;
  }

  static XmlElement _insertCell(XmlElement row, XmlElement? lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var cell = _createCell(columnIndex, rowIndex, value);
    if (lastCell == null) {
      row.children.add(cell);
    } else {
      var index = row.children.indexOf(lastCell);
      row.children.insert(index, cell);
    }
    return cell;
  }

  static XmlElement _replaceCell(XmlElement row, XmlElement? lastCell,
      int columnIndex, int rowIndex, dynamic value) {
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
      XmlAttribute(
          XmlName('r'), '${numericToLetters(columnIndex + 1)}${rowIndex + 1}'),
      XmlAttribute(XmlName('t'), 'inlineStr'),
    ];
    var children = value == null
        ? <XmlElement>[]
        : <XmlElement>[
            XmlElement(XmlName('is'), [], [
              XmlElement(XmlName('t'), [], [XmlText(value.toString())])
            ]),
          ];
    return XmlElement(XmlName('c'), attributes, children);
  }
}
