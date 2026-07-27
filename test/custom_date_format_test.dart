@TestOn('vm')
library spreadsheet_custom_date_format_test;

import 'package:test/test.dart';

import 'common_io.dart';

/// Tests that custom <numFmt> date format codes (numFmtId outside the
/// built-in 14-22/45-47 range) are recognised and decoded as formatted
/// date strings instead of leaking through as raw numbers.
void main() {
  group('Custom date number format (xlsx):', () {
    test('sample_date_format.xlsx — dd/mm/yyyy custom numFmt', () {
      var decoder = decode('sample_date_format.xlsx');
      var sheet = decoder.tables.values.first;

      // Cells L2, P2, R2 use the custom numFmtId=59 with
      // formatCode="dd/mm/yyyy" and raw serial values 39650, 45782, 46144.
      // Column L = index 11, P = 15, R = 17 (0-based).
      final row2 = sheet.rows[1];

      final ddmmyyyy = RegExp(r'^\d{2}/\d{2}/\d{4}$');

      expect(row2[11], isA<String>(),
          reason: 'L2 must decode as a date string, not a number');
      expect(row2[15], isA<String>(),
          reason: 'P2 must decode as a date string, not a number');
      expect(row2[17], isA<String>(),
          reason: 'R2 must decode as a date string, not a number');

      expect(row2[11] as String, matches(ddmmyyyy),
          reason: 'L2 should match the dd/mm/yyyy format code');
      expect(row2[15] as String, matches(ddmmyyyy));
      expect(row2[17] as String, matches(ddmmyyyy));

      // The exact day-of-month follows the existing decoder's serial-to-date
      // math (1899-12-30 baseline, the same convention used for the
      // built-in numFmtId 14-22 path). These are the deterministic outputs
      // for the three serials in this fixture.
      expect(row2[11], equals('21/07/2008'));
      expect(row2[15], equals('05/05/2025'));
      expect(row2[17], equals('02/05/2026'));

      // Sanity: none of the date cells should leak through as numbers.
      expect(row2[11], isNot(isA<num>()));
      expect(row2[15], isNot(isA<num>()));
      expect(row2[17], isNot(isA<num>()));
    });
  });

  group('Built-in date numFmt + dateFormat option (xlsx):', () {
    test('default dateFormat renders built-in date as yyyy-MM-dd', () {
      // new_sample_file_bytco.xlsx contains a column whose cell uses the
      // built-in numFmtId=14 (m/d/yyyy). The decoded value must be a
      // date-only string in the package default format, not an ISO
      // timestamp like 2008-07-21T00:00:00.000.
      var decoder = decode('new_sample_file_bytco.xlsx');
      var sheet = decoder.tables.values.first;

      final iso = RegExp(r'^\d{4}-\d{2}-\d{2}$');
      var foundDateCell = false;
      for (var row in sheet.rows) {
        for (var cell in row) {
          if (cell is String && iso.hasMatch(cell)) {
            foundDateCell = true;
            // No leftover ISO time component.
            expect(cell, isNot(contains('T')));
            expect(cell, isNot(contains(':')));
          }
        }
      }
      expect(foundDateCell, isTrue,
          reason: 'Expected at least one yyyy-MM-dd date cell in fixture');
    });

    test('caller-supplied dateFormat dd/MM/yyyy is honoured', () {
      var decoder =
          decode('new_sample_file_bytco.xlsx', dateFormat: 'dd/MM/yyyy');
      var sheet = decoder.tables.values.first;

      final ddmmyyyy = RegExp(r'^\d{2}/\d{2}/\d{4}$');
      var foundDateCell = false;
      for (var row in sheet.rows) {
        for (var cell in row) {
          if (cell is String && ddmmyyyy.hasMatch(cell)) {
            foundDateCell = true;
          }
        }
      }
      expect(foundDateCell, isTrue,
          reason: 'Expected at least one dd/MM/yyyy date cell');
    });

    test('caller-supplied dateFormat overrides custom workbook code', () {
      // sample_date_format.xlsx normally renders custom-numFmt dates as
      // dd/mm/yyyy because that's the workbook's format code. When the
      // caller passes a non-default dateFormat the custom branch must
      // honour the caller's preference instead.
      var decoder = decode('sample_date_format.xlsx', dateFormat: 'yyyy.MM.dd');
      var sheet = decoder.tables.values.first;
      final row2 = sheet.rows[1];
      expect(row2[11], equals('2008.07.21'));
      expect(row2[15], equals('2025.05.05'));
      expect(row2[17], equals('2026.05.02'));
    });
  });
}
