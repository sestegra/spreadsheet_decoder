library ods_test;

import 'package:spreadsheet/spreadsheet.dart';
import 'package:test/test.dart';

import 'common.dart';

void main() {
  group('ods spreadsheet', () {
    test('Check format', () {
      var decoder = decode('default.ods');
      expect(decoder is OdsDecoder, isTrue);
    });
  });
}
