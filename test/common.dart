part of spreadsheet_test;

var expectedTest = <String, List<List>>{
  'ONE': [
    ['A', 'B', 'C'],
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
    [12, 15, 18],
    [3, 3, 3],
    [3, 3, 3],
    [3, 3, 3],
    [null, null, null],
    [6, 7, 8],
    [6, 7, 8],
    [6, 7, 8]
  ],
  'TWO': [
    ['X', 'Y', 'Z'],
    [10, 11, 12],
    [13, null, 15],
    [16, 17, 18],
    ["&é'(§è!çà)-", ' "«A»"', '<>']
  ],
  'THREE': [
    ['P', 'Q', 'R'],
    [100, 101, 102],
    [103, 104, 105],
    [106, 107, 108],
    ['A', 'B\nC', 'D\nE\nF']
  ],
  'EMPTY': []
};

var expectedPerl = <String, List<List>>{
  'Sheet1': [
    [-1500.99, 17, null],
    [null, null, null],
    ['one', 'more', 'cell']
  ],
  'Sheet2': [],
  'Sheet3': [
    ['Both alike', 'Both alike', null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, 'Cell C14']
  ]
};

var expectedPerlAfterUpdate = <String, List<List>>{
  'Sheet1': [
    ['0x0', '1x0', '2x0'],
    ['0x1', '1x1', '2x1'],
    ['0x2', '1x2', '2x2'],
  ],
  'Sheet2': [],
  'Sheet3': [
    ['Both alike', 'Both alike', null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, 'Cell B6', null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, null],
    [null, null, 'Cell C14']
  ]
};

var expectedFormat = <String, List<List>>{
  'Sheet1': [
    [
      1337,
      1337.42,
      1234.56,
      true,
      'Hello World',
      'One line\nTwo lines\nThree lines',
      0.124,
      '2006-02-01T00:00:00.000',
      '2006-02-01T13:37:00.000',
      '2006-02-01T13:37:42.000',
      '13:37:00',
      '13:37:42',
    ]
  ]
};

var expectedEmptyColumn = <String, List<List>>{
  'EMPTY': [
    ['a1', null, 'c1'],
    ['a2', null, 'c2'],
    ['a3', null, 'c3']
  ]
};

var expectNumbers = <String, List<List>>{
  'Blatt 1': [
    ['Tabelle 1']
  ]
};

Map<String, List<List>> copyTables(Map<String, List<List>> tables) {
  var copy = <String, List<List>>{};
  tables.forEach((sheet, table) {
    copy[sheet] = List<List>.from(table.map((row) => List.from(row)));
  });
  return copy;
}

void testUnsupported() {
  test('Unsupported file', () {
    SpreadsheetDecoder? decoder;
    try {
      decoder = decode('unsupported.zip');
    } catch (e) {
      expect(e, isUnsupportedError);
    }
    expect(decoder, isNull);
  });
}

void testXlsx() {
  group('lettersToNumeric:', () {
    test('Simple capital letter', () {
      expect(lettersToNumeric('A'), 1);
      expect(lettersToNumeric('B'), 2);
      expect(lettersToNumeric('C'), 3);
      expect(lettersToNumeric('D'), 4);
      expect(lettersToNumeric('E'), 5);
      expect(lettersToNumeric('F'), 6);
      expect(lettersToNumeric('G'), 7);
      expect(lettersToNumeric('H'), 8);
      expect(lettersToNumeric('I'), 9);
      expect(lettersToNumeric('J'), 10);
      expect(lettersToNumeric('K'), 11);
      expect(lettersToNumeric('L'), 12);
      expect(lettersToNumeric('M'), 13);
      expect(lettersToNumeric('N'), 14);
      expect(lettersToNumeric('O'), 15);
      expect(lettersToNumeric('P'), 16);
      expect(lettersToNumeric('Q'), 17);
      expect(lettersToNumeric('R'), 18);
      expect(lettersToNumeric('S'), 19);
      expect(lettersToNumeric('T'), 20);
      expect(lettersToNumeric('U'), 21);
      expect(lettersToNumeric('V'), 22);
      expect(lettersToNumeric('W'), 23);
      expect(lettersToNumeric('X'), 24);
      expect(lettersToNumeric('Y'), 25);
      expect(lettersToNumeric('Z'), 26);
    });

    test('Simple lower letter', () {
      expect(lettersToNumeric('a'), 1);
      expect(lettersToNumeric('b'), 2);
      expect(lettersToNumeric('c'), 3);
      expect(lettersToNumeric('d'), 4);
      expect(lettersToNumeric('e'), 5);
      expect(lettersToNumeric('f'), 6);
      expect(lettersToNumeric('g'), 7);
      expect(lettersToNumeric('h'), 8);
      expect(lettersToNumeric('i'), 9);
      expect(lettersToNumeric('j'), 10);
      expect(lettersToNumeric('k'), 11);
      expect(lettersToNumeric('l'), 12);
      expect(lettersToNumeric('m'), 13);
      expect(lettersToNumeric('n'), 14);
      expect(lettersToNumeric('o'), 15);
      expect(lettersToNumeric('p'), 16);
      expect(lettersToNumeric('q'), 17);
      expect(lettersToNumeric('r'), 18);
      expect(lettersToNumeric('s'), 19);
      expect(lettersToNumeric('t'), 20);
      expect(lettersToNumeric('u'), 21);
      expect(lettersToNumeric('v'), 22);
      expect(lettersToNumeric('w'), 23);
      expect(lettersToNumeric('x'), 24);
      expect(lettersToNumeric('y'), 25);
      expect(lettersToNumeric('z'), 26);
    });

    test('Up to AMJ', () {
      expect(lettersToNumeric('amj'), 1024);
      expect(lettersToNumeric('Amj'), 1024);
      expect(lettersToNumeric('AMj'), 1024);
      expect(lettersToNumeric('AMJ'), 1024);
      expect(lettersToNumeric('aMj'), 1024);
      expect(lettersToNumeric('aMJ'), 1024);
    });
  });

  group('numericToLetters:', () {
    test('Simple capital letter', () {
      expect(numericToLetters(1), 'A');
      expect(numericToLetters(2), 'B');
      expect(numericToLetters(3), 'C');
      expect(numericToLetters(4), 'D');
      expect(numericToLetters(5), 'E');
      expect(numericToLetters(6), 'F');
      expect(numericToLetters(7), 'G');
      expect(numericToLetters(8), 'H');
      expect(numericToLetters(9), 'I');
      expect(numericToLetters(10), 'J');
      expect(numericToLetters(11), 'K');
      expect(numericToLetters(12), 'L');
      expect(numericToLetters(13), 'M');
      expect(numericToLetters(14), 'N');
      expect(numericToLetters(15), 'O');
      expect(numericToLetters(16), 'P');
      expect(numericToLetters(17), 'Q');
      expect(numericToLetters(18), 'R');
      expect(numericToLetters(19), 'S');
      expect(numericToLetters(20), 'T');
      expect(numericToLetters(21), 'U');
      expect(numericToLetters(22), 'V');
      expect(numericToLetters(23), 'W');
      expect(numericToLetters(24), 'X');
      expect(numericToLetters(25), 'Y');
      expect(numericToLetters(26), 'Z');
    });

    test('Up to AMJ', () {
      expect(numericToLetters(1024), 'AMJ');
    });
  });

  group('cellCoordsFromCellId:', () {
    test('Simple coords', () {
      expect(cellCoordsFromCellId('A1'), [1, 1]);
      expect(cellCoordsFromCellId('B3'), [2, 3]);
      expect(cellCoordsFromCellId('AMJ42'), [1024, 42]);
      expect(cellCoordsFromCellId('aMj1337'), [1024, 1337]);
    });
  });

  group('xlsx spreadsheet:', () {
    test('Check format', () {
      var decoder = decode('default.xlsx');
      expect(decoder is XlsxDecoder, isTrue);
    });

    test('Empty file', () {
      var decoder = decode('default.xlsx');
      expect(decoder.tables.length, 3);
      expect(decoder.tables['Sheet1']!.rows, []);
      expect(decoder.tables['Sheet2']!.rows, []);
      expect(decoder.tables['Sheet3']!.rows, []);
    });

    test('Test file', () {
      var decoder = decode('test.xlsx');
      expect(decoder.tables.length, expectedTest.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedTest[name]);
      });
    });

    test('Perl file', () {
      var decoder = decode('perl.xlsx');
      expect(decoder.tables.length, expectedPerl.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedPerl[name]);
      });
    });

    test('Format file:', () {
      var decoder = decode('format.xlsx');
      expect(decoder.tables.length, expectedFormat.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedFormat[name]);
      });
    });

    test('Empty column file:', () {
      var decoder = decode('empty_column.xlsx');
      expect(decoder.tables.length, expectedEmptyColumn.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedEmptyColumn[name]);
      });
    });

    test('Number to xlsx:', () {
      var decoder = decode('numbers_to_xlsx.xlsx');
      expect(decoder.tables.length, expectNumbers.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectNumbers[name]);
      });
    });

    test('Check `Target` format with absolute path', () {
      var decoder = decode('absolute_target.xlsx');
      expect(decoder is XlsxDecoder, isTrue);
    }, onPlatform: {
      'browser': Skip('Not supported in browser'),
    });
  });
}

void testOds() {
  group('ods spreadsheet:', () {
    test('Check format', () {
      var decoder = decode('default.ods');
      expect(decoder is OdsDecoder, isTrue);
    });

    test('Empty file', () {
      var decoder = decode('default.ods');
      expect(decoder.tables.length, 1);
      expect(decoder.tables['Sheet1']!.rows, []);
    });

    test('Test file', () {
      var decoder = decode('test.ods');
      expect(decoder.tables.length, expectedTest.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedTest[name]);
      });
    });

    test('Perl file', () {
      var decoder = decode('perl.ods');
      expect(decoder.tables.length, expectedPerl.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedPerl[name]);
      });
    });

    test('Format file:', () {
      var decoder = decode('format.ods');
      expect(decoder.tables.length, expectedFormat.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedFormat[name]);
      });
    });

    test('Empty column file:', () {
      var decoder = decode('empty_column.ods');
      expect(decoder.tables.length, expectedEmptyColumn.keys.length);
      decoder.tables.forEach((name, table) {
        expect(table.rows, expectedEmptyColumn[name]);
      });
    });
  });
}

void testUpdateXlsx() {
  group('xlsx spreadsheet update:', () {
    group('No update', () {
      test('Empty file', () {
        var first = decode('default.xlsx', update: true);
        var data = first.encode();
        save('test/out/no/default.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, 3);
        expect(decoder.tables['Sheet1']!.rows, []);
        expect(decoder.tables['Sheet2']!.rows, []);
        expect(decoder.tables['Sheet3']!.rows, []);
      });

      test('Test file', () {
        var first = decode('test.xlsx', update: true);
        var data = first.encode();
        save('test/out/no/test.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedTest.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedTest[name]);
        });
      });

      test('Perl file', () {
        var first = decode('perl.xlsx', update: true);
        var data = first.encode();
        save('test/out/no/perl.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerl.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerl[name]);
        });
      });

      test('Format file:', () {
        var first = decode('format.xlsx', update: true);
        var data = first.encode();
        save('test/out/no/format.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedFormat.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedFormat[name]);
        });
      });

      test('Empty column file:', () {
        var first = decode('empty_column.xlsx', update: true);
        var data = first.encode();
        save('test/out/no/empty_column.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedEmptyColumn.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedEmptyColumn[name]);
        });
      });

      test('dataUrl', () {
        var xlsx = decode('test.xlsx', update: true);
        expect(xlsx.extension, '.xlsx');
        var data = xlsx.dataUrl();

        var decoder =
            SpreadsheetDecoder.decodeBytes(base64Decode(data.split(',').last));
        expect(decoder.tables.length, expectedTest.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedTest[name]);
        });
      });
    });

    group('Update', () {
      test('Perl file #1', () {
        var first = decode('perl.xlsx', update: true)
          ..updateCell('Sheet1', 0, 0, '0x0')
          ..updateCell('Sheet1', 1, 0, '1x0')
          ..updateCell('Sheet1', 2, 0, '2x0')
          ..updateCell('Sheet1', 0, 1, '0x1')
          ..updateCell('Sheet1', 1, 1, '1x1')
          ..updateCell('Sheet1', 2, 1, '2x1')
          ..updateCell('Sheet1', 0, 2, '0x2')
          ..updateCell('Sheet1', 1, 2, '1x2')
          ..updateCell('Sheet1', 2, 2, '2x2')
          ..updateCell('Sheet3', 1, 5, 'Cell B6');
        var data = first.encode();
        save('test/out/update/perl.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerlAfterUpdate.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerlAfterUpdate[name]);
        });
      });

      test('Perl file #2', () {
        var first = decode('perl.xlsx', update: true)
          ..updateCell('Sheet1', 2, 0, '2x0')
          ..updateCell('Sheet1', 1, 0, '1x0')
          ..updateCell('Sheet1', 0, 0, '0x0')
          ..updateCell('Sheet1', 2, 1, '2x1')
          ..updateCell('Sheet1', 1, 1, '1x1')
          ..updateCell('Sheet1', 0, 1, '0x1')
          ..updateCell('Sheet1', 2, 2, '2x2')
          ..updateCell('Sheet1', 1, 2, '1x2')
          ..updateCell('Sheet1', 0, 2, '0x2')
          ..updateCell('Sheet3', 1, 5, 'Cell B6');
        var data = first.encode();
        save('test/out/update/perl.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerlAfterUpdate.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerlAfterUpdate[name]);
        });
      });

      test('Perl file #3', () {
        var first = decode('perl.xlsx', update: true)
          ..updateCell('Sheet1', 1, 0, '1x0')
          ..updateCell('Sheet1', 2, 0, '2x0')
          ..updateCell('Sheet1', 0, 0, '0x0')
          ..updateCell('Sheet1', 1, 1, '1x1')
          ..updateCell('Sheet1', 2, 1, '2x1')
          ..updateCell('Sheet1', 0, 1, '0x1')
          ..updateCell('Sheet1', 1, 2, '1x2')
          ..updateCell('Sheet1', 2, 2, '2x2')
          ..updateCell('Sheet1', 0, 2, '0x2')
          ..updateCell('Sheet3', 1, 5, 'Cell B6');
        var data = first.encode();
        save('test/out/update/perl.xlsx', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerlAfterUpdate.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerlAfterUpdate[name]);
        });
      });
    });

    group('row', () {
      group('insert', () {
        test('at top', () {
          var first = decode('test.xlsx', update: true)..insertRow('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.insert(0, [null, null, null]);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at bottom', () {
          var first = decode('test.xlsx', update: true)
            ..insertRow('ONE', 12)
            ..updateCell('ONE', 0, 12, 'insert');
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.add(['insert', null, null]);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.xlsx', update: true)..insertRow('ONE', 10);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.insert(10, [null, null, null]);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in empty sheet', () {
          var first = decode('test.xlsx', update: true)
            ..insertRow('EMPTY', 0)
            ..insertColumn('EMPTY', 0)
            ..updateCell('EMPTY', 0, 0, '42');
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['EMPTY']!.insert(0, ['42']);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.xlsx')..insertRow('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () => decode('test.xlsx', update: true)..insertRow('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(() => decode('test.xlsx', update: true)..insertRow('ONE', 13),
              throwsRangeError);
        });
      });

      group('remove', () {
        test('at top', () {
          var first = decode('test.xlsx', update: true)..removeRow('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.removeAt(0);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at bottom', () {
          var first = decode('test.xlsx', update: true)..removeRow('ONE', 11);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.removeAt(11);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.xlsx', update: true)..removeRow('ONE', 10);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.removeAt(10);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.xlsx')..removeRow('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () => decode('test.xlsx', update: true)..removeRow('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(() => decode('test.xlsx', update: true)..removeRow('ONE', 13),
              throwsRangeError);
        });
      });
    });

    group('col', () {
      group('insert', () {
        test('at left', () {
          var first = decode('test.xlsx', update: true)..insertColumn('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.insert(0, null);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at right', () {
          var first = decode('test.xlsx', update: true)
            ..insertColumn('ONE', 3)
            ..updateCell('ONE', 3, 0, 'insert');
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.add(null);
          }
          expectedTables['ONE']![0][3] = 'insert';

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.xlsx', update: true)..insertColumn('ONE', 1);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.insert(1, null);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.xlsx')..insertColumn('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () =>
                  decode('test.xlsx', update: true)..insertColumn('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(
              () => decode('test.xlsx', update: true)..insertColumn('ONE', 13),
              throwsRangeError);
        });
      });

      group('remove', () {
        test('at left', () {
          var first = decode('test.xlsx', update: true)..removeColumn('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.removeAt(0);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at right', () {
          var first = decode('test.xlsx', update: true)..removeColumn('ONE', 2);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.removeAt(2);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.xlsx', update: true)..removeColumn('ONE', 1);
          var data = first.encode();
          save('test/out/update/test.xlsx', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.removeAt(1);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.xlsx')..removeColumn('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () =>
                  decode('test.xlsx', update: true)..removeColumn('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(
              () => decode('test.xlsx', update: true)..removeColumn('ONE', 13),
              throwsRangeError);
        });
      });
    });
  });
}

void testUpdateOds() {
  group('ods spreadsheet update:', () {
    group('No update', () {
      test('Empty file', () {
        var first = decode('default.ods', update: true);
        var data = first.encode();
        save('test/out/no/default.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, 1);
        expect(decoder.tables['Sheet1']!.rows, []);
      });

      test('Test file', () {
        var first = decode('test.ods', update: true);
        var data = first.encode();
        save('test/out/no/test.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedTest.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedTest[name]);
        });
      });

      test('Perl file', () {
        var first = decode('perl.ods', update: true);
        var data = first.encode();
        save('test/out/no/perl.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerl.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerl[name]);
        });
      });

      test('Format file:', () {
        var first = decode('format.ods', update: true);
        var data = first.encode();
        save('test/out/no/format.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedFormat.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedFormat[name]);
        });
      });

      test('Empty column file:', () {
        var first = decode('empty_column.ods', update: true);
        var data = first.encode();
        save('test/out/no/empty_column.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedEmptyColumn.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedEmptyColumn[name]);
        });
      });

      test('dataUrl', () {
        var ods = decode('test.ods', update: true);
        expect(ods.extension, '.ods');
        var data = ods.dataUrl();

        var decoder =
            SpreadsheetDecoder.decodeBytes(base64Decode(data.split(',').last));
        expect(decoder.tables.length, expectedTest.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedTest[name]);
        });
      });
    });

    group('Update', () {
      test('Perl file #1', () {
        var first = decode('perl.ods', update: true)
          ..updateCell('Sheet1', 0, 0, '0x0')
          ..updateCell('Sheet1', 1, 0, '1x0')
          ..updateCell('Sheet1', 2, 0, '2x0')
          ..updateCell('Sheet1', 0, 1, '0x1')
          ..updateCell('Sheet1', 1, 1, '1x1')
          ..updateCell('Sheet1', 2, 1, '2x1')
          ..updateCell('Sheet1', 0, 2, '0x2')
          ..updateCell('Sheet1', 1, 2, '1x2')
          ..updateCell('Sheet1', 2, 2, '2x2')
          ..updateCell('Sheet3', 1, 5, 'Cell B6');
        var data = first.encode();
        save('test/out/update/perl.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerlAfterUpdate.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerlAfterUpdate[name]);
        });
      });

      test('Perl file #2', () {
        var first = decode('perl.ods', update: true)
          ..updateCell('Sheet1', 2, 0, '2x0')
          ..updateCell('Sheet1', 1, 0, '1x0')
          ..updateCell('Sheet1', 0, 0, '0x0')
          ..updateCell('Sheet1', 2, 1, '2x1')
          ..updateCell('Sheet1', 1, 1, '1x1')
          ..updateCell('Sheet1', 0, 1, '0x1')
          ..updateCell('Sheet1', 2, 2, '2x2')
          ..updateCell('Sheet1', 1, 2, '1x2')
          ..updateCell('Sheet1', 0, 2, '0x2')
          ..updateCell('Sheet3', 1, 5, 'Cell B6');
        var data = first.encode();
        save('test/out/update/perl.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerlAfterUpdate.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerlAfterUpdate[name]);
        });
      });

      test('Perl file #3', () {
        var first = decode('perl.ods', update: true)
          ..updateCell('Sheet1', 1, 0, '1x0')
          ..updateCell('Sheet1', 2, 0, '2x0')
          ..updateCell('Sheet1', 0, 0, '0x0')
          ..updateCell('Sheet1', 1, 1, '1x1')
          ..updateCell('Sheet1', 2, 1, '2x1')
          ..updateCell('Sheet1', 0, 1, '0x1')
          ..updateCell('Sheet1', 1, 2, '1x2')
          ..updateCell('Sheet1', 2, 2, '2x2')
          ..updateCell('Sheet1', 0, 2, '0x2')
          ..updateCell('Sheet3', 1, 5, 'Cell B6');
        var data = first.encode();
        save('test/out/update/perl.ods', data);

        var decoder = SpreadsheetDecoder.decodeBytes(data);
        expect(decoder.tables.length, expectedPerlAfterUpdate.keys.length);
        decoder.tables.forEach((name, table) {
          expect(table.rows, expectedPerlAfterUpdate[name]);
        });
      });
    });

    group('row', () {
      group('insert', () {
        test('at top', () {
          var first = decode('test.ods', update: true)..insertRow('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.insert(0, [null, null, null]);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at bottom', () {
          var first = decode('test.ods', update: true)
            ..insertRow('ONE', 12)
            ..updateCell('ONE', 0, 12, 'insert');
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.add(['insert', null, null]);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.ods', update: true)..insertRow('ONE', 10);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.insert(10, [null, null, null]);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.ods')..insertRow('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () => decode('test.ods', update: true)..insertRow('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(() => decode('test.ods', update: true)..insertRow('ONE', 13),
              throwsRangeError);
        });
      });

      group('remove', () {
        test('at top', () {
          var first = decode('test.ods', update: true)..removeRow('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.removeAt(0);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at bottom', () {
          var first = decode('test.ods', update: true)..removeRow('ONE', 11);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.removeAt(11);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.ods', update: true)..removeRow('ONE', 10);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          expectedTables['ONE']!.removeAt(10);

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.ods')..removeRow('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () => decode('test.ods', update: true)..removeRow('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(() => decode('test.ods', update: true)..removeRow('ONE', 13),
              throwsRangeError);
        });
      });
    });

    group('col', () {
      group('insert', () {
        test('at left', () {
          var first = decode('test.ods', update: true)..insertColumn('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.insert(0, null);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at right', () {
          var first = decode('test.ods', update: true)
            ..insertColumn('ONE', 3)
            ..updateCell('ONE', 3, 0, 'insert');
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.add(null);
          }
          expectedTables['ONE']![0][3] = 'insert';

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.ods', update: true)..insertColumn('ONE', 1);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.insert(1, null);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.ods')..insertColumn('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () =>
                  decode('test.ods', update: true)..insertColumn('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(
              () => decode('test.ods', update: true)..insertColumn('ONE', 13),
              throwsRangeError);
        });
      });

      group('remove', () {
        test('at left', () {
          var first = decode('test.ods', update: true)..removeColumn('ONE', 0);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.removeAt(0);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('at right', () {
          var first = decode('test.ods', update: true)..removeColumn('ONE', 2);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.removeAt(2);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('in between', () {
          var first = decode('test.ods', update: true)..removeColumn('ONE', 1);
          var data = first.encode();
          save('test/out/update/test.ods', data);

          var decoder = SpreadsheetDecoder.decodeBytes(data);
          var expectedTables = copyTables(expectedTest);
          for (var row in expectedTables['ONE']!) {
            row.removeAt(1);
          }

          expect(decoder.tables.length, expectedTables.keys.length);
          decoder.tables.forEach((name, table) {
            expect(table.rows, expectedTables[name]);
          });
        });

        test('ArgumentError exception', () {
          expect(() => decode('test.ods')..removeColumn('ONE', 1),
              throwsArgumentError);
        });

        test('ArgumentError exception', () {
          expect(
              () =>
                  decode('test.ods', update: true)..removeColumn('UNKNOWN', 1),
              throwsArgumentError);
        });

        test('RangeError exception', () {
          expect(
              () => decode('test.ods', update: true)..removeColumn('ONE', 13),
              throwsRangeError);
        });
      });
    });
  });
}
