part of spreadsheet_test;

var expectedTest = {
  "ONE": [
    [  "A",   "B",   "C"],
    [  "1",   "2",   "3"],
    [  "4",   "5",   "6"],
    [  "7",   "8",   "9"],
    [ "12",  "15",  "18"],
    [  "3",   "3",   "3"],
    [  "3",   "3",   "3"],
    [  "3",   "3",   "3"],
    [ null,  null,  null],
    [  "6",   "7",   "8"],
    [  "6",   "7",   "8"],
    [  "6",   "7",   "8"]
  ],
  "TWO": [
    [  "X",   "Y",   "Z"],
    [ "10",  "11",  "12"],
    [ "13",  "14",  "15"],
    [ "16",  "17",  "18"],
    ["&é'(§è!çà)-", '"«A»"', "<>"]
  ],
  "THREE": [
    [  "P",   "Q",   "R"],
    ["100", "101", "102"],
    ["103", "104", "105"],
    ["106", "107", "108"],
    ["A", r"B\nC", r"D\nE\nF"]],
  "EMPTY": []
};

var expectedPerl = {
  "Sheet1": [
    ['-\$1,500.99', '17', null],
    [null, null, null],
    ['one', 'more', 'cell']
  ],
  "Sheet2": [],
  "Sheet3": [
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

testXlsx() {
  group('xlsx spreadsheet', () {
    test('Check format', () {
      var decoder = decode('default.xlsx');
      expect(decoder is XlsxDecoder, isTrue);
    });

    test('Empty file', () {
      var decoder = decode('default.xlsx');
      expect(decoder.tables.length, 3);
    }, skip: 'Not yet implemented');
  });
}

testOds() {
  group('ods spreadsheet', () {
    test('Check format', () {
      var decoder = decode('default.ods');
      expect(decoder is OdsDecoder, isTrue);
    });

    test('Empty file', () {
      var decoder = decode('default.ods');
      expect(decoder.tables.length, 1);
      expect(decoder.tables['Sheet1'].rows, []);
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
  });
}
