@TestOn('vm')

library spreadsheet_test;

import 'dart:convert';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';
import 'package:test/test.dart';

import 'common_io.dart';
part 'common.dart';

void main() {
  testUnsupported();
  testOds();
  testXlsx();
  testUpdateOds();
  testUpdateXlsx();
}
