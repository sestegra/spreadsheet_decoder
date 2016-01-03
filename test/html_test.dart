@TestOn('browser')

library spreadsheet_test;

import 'package:spreadsheet/spreadsheet.dart';
import 'package:test/test.dart';

import 'common_html.dart';
part 'common.dart';

void main() {
  testOds();
  testXlsx();
}
