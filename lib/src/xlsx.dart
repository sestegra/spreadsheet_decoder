part of spreadsheet_decoder;

const String _relationships =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const String _relationshipsStyles =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const String _relationshipsWorksheet =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const String _relationshipsSharedStrings =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';

extension XlsxXmlFindExtension on XmlNode {
  Iterable<XmlElement> findElementsStar(String name,
          {String namespace = "*"}) =>
      findElements(name, namespace: namespace);

  Iterable<XmlElement> findAllElementsStar(String name,
          {String namespace = "*"}) =>
      findAllElements(name, namespace: namespace);
}

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
  String? _stylesTarget;
  String? _sharedStringsTarget;
  final Map<String, String> _worksheetTargets = <String, String>{};

  XlsxDecoder(Archive archive, {bool update = false}) {
    _archive = archive;
    _update = update;
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
      document.findAllElementsStar('Relationship').forEach((node) {
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
    if (_stylesTarget is String) {
      final namePath = _stylesTarget!.startsWith('/')
          ? _stylesTarget!.substring(1)
          : 'xl/$_stylesTarget';
      var styles = _archive.findFile(namePath);
      if (styles != null) {
        styles.decompress();
        var document = XmlDocument.parse(utf8.decode(styles.content));
        document
            .findAllElementsStar('cellXfs')
            .first
            .findElementsStar('xf')
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
  }

  void _parseSharedStrings() {
    if (_sharedStringsTarget is String) {
      final namePath = _sharedStringsTarget!.startsWith('/')
          ? _sharedStringsTarget!.substring(1)
          : 'xl/$_sharedStringsTarget';
      var sharedStrings = _archive.findFile(namePath);
      if (sharedStrings != null) {
        sharedStrings.decompress();
        var document = XmlDocument.parse(utf8.decode(sharedStrings.content));
        document.findAllElementsStar('si').forEach((node) {
          _parseSharedString(node);
        });
      }
    }
  }

  String _parseRichText(XmlElement node) {
    return _parseValue(node.findElementsStar('t').first);
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
    workbook?.decompress();
    var document = XmlDocument.parse(utf8.decode(workbook?.content));
    document.findAllElementsStar('sheet').forEach((node) {
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
    file?.decompress();

    var content = XmlDocument.parse(utf8.decode(file?.content));
    var worksheet = content.findElementsStar('worksheet').first;
    var sheet = worksheet.findElementsStar('sheetData').first;

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
            int.parse(_parseValue(node.findElementsStar('v').first))];
        break;
      // boolean
      case 'b':
        value = _parseValue(node.findElementsStar('v').first) == '1';
        break;
      // error
      case 'e':
      // formula
      case 'str':
        // <c r="C6" s="1" vm="15" t="str">
        //  <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
        //  <v>2838512.355</v>
        // </c>
        value = _parseValue(node.findElementsStar('v').first);
        break;
      // inline string
      case 'inlineStr':
        // <c r="B2" t="inlineStr">
        // <is><t>Hello world</t></is>
        // </c>
        value = _parseValue(node.findAllElementsStar('t').first);
        break;
      // number
      case 'n':
      default:
        var s = node.getAttribute('s');
        var valueNode = node.findElementsStar('v');
        var content = valueNode.first;
        if (s != null) {
          var fmtId = _numFormats[int.parse(s)];
          // date
          if (((fmtId >= 14) && (fmtId <= 17)) || (fmtId == 22)) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(1899, 12, 30);
            value = date
                .add(Duration(milliseconds: delta.toInt()))
                .toIso8601String();
            // time
          } else if (((fmtId >= 18) && (fmtId <= 21)) ||
              ((fmtId >= 45) && (fmtId <= 47))) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(0);
            date = date.add(Duration(milliseconds: delta.toInt()));
            value =
                '${_twoDigits(date.hour)}:${_twoDigits(date.minute)}:${_twoDigits(date.second)}';
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
      table.findElementsStar('row');

  static Iterable<XmlElement> _findCells(XmlElement row) =>
      row.findElementsStar('c');

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
