library xlsx_test;

import 'dart:io';
import 'package:spreadsheet/spreadsheet.dart';
import 'package:test/test.dart';

import 'common.dart';

void main() {
  group('xlsx spreadsheet', () {
    test('Check format', () {
      var decoder = decode('default.xlsx');
      expect(decoder is XlsxDecoder, isTrue);
    });
  });
}
