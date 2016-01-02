part of spreadsheet_test;

testXlsx() {
  group('xlsx spreadsheet', () {
    test('Check format', () async {
      var decoder = await decode('default.xlsx');
      expect(decoder is XlsxDecoder, isTrue);
    });
  });
}

testOds() {
  group('ods spreadsheet', () {
    test('Check format', () async {
      var decoder = await decode('default.ods');
      expect(decoder is OdsDecoder, isTrue);
    });
  });
}
