@TestOn('vm')
library spreadsheet_teaching_plan_sample_test;

import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';
import 'package:test/test.dart';

import 'common_io.dart';

const filename =
    'subject-Data Structures and Algorithms-of-FY - MCA-teaching-plan-sample-1784613441518.xlsx';
const sheetName = 'Teaching Plan Sample';

const headers = [
  'Topic',
  'SubTopic',
  'CourseOutcome',
  'LearningOutcome',
  'Hours',
  'Minutes',
  'PlanningDate',
  'ExecutionDate',
];

const topicCol = 0;
const subTopicCol = 1;
const courseOutcomeCol = 2;
const learningOutcomeCol = 3;
const hoursCol = 4;
const minutesCol = 5;
const planningDateCol = 6;
const executionDateCol = 7;

/// The PlanningDate column mixes two storage forms in the source workbook:
///
///  * a real date serial styled with `numFmtId="58"` (a built-in
///    locale-specific date format) — these are the rows that used to decode
///    as a bare number such as `46212`;
///  * a shared string holding pre-formatted `d/M/yyyy` text — passed through
///    verbatim by the decoder.
///
/// Keyed by row index in [SpreadsheetTable.rows] (row 0 is the header).
const serialBackedPlanningDates = <int, String>{
  1: '2026-07-09', // serial 46212
  2: '2026-09-09', // serial 46274
  3: '2026-11-09', // serial 46335
  13: '2026-05-10', // serial 46152
  14: '2026-07-10', // serial 46213
  15: '2026-09-10', // serial 46275
  16: '2026-12-10', // serial 46366
  25: '2026-02-11', // serial 46064
  26: '2026-04-11', // serial 46123
  35: '2026-02-12', // serial 46065
  36: '2026-04-12', // serial 46124
  37: '2026-07-12', // serial 46215
  38: '2026-09-12', // serial 46277
  39: '2026-11-12', // serial 46338
};

const textBackedPlanningDates = <int, String>{
  4: '16/9/2026',
  5: '18/9/2026',
  6: '18/9/2026',
  7: '21/9/2026',
  8: '21/9/2026',
  9: '21/9/2026',
  10: '23/9/2026',
  11: '28/9/2026',
  12: '30/9/2026',
  17: '14/10/2026',
  18: '16/10/2026',
  19: '19/10/2026',
  20: '21/10/2026',
  21: '23/10/2026',
  22: '26/10/2026',
  23: '28/10/2026',
  24: '30/10/2026',
  27: '13/11/2026',
  28: '16/11/2026',
  29: '18/11/2026',
  30: '20/11/2026',
  31: '23/11/2026',
  32: '25/11/2026',
  33: '27/11/2026',
  34: '30/11/2026',
  40: '14/12/2026',
  41: '16/12/2026',
};

String _cell(dynamic value, {int width = 34}) {
  if (value == null) return 'null';
  var text = '${value.runtimeType}:${value.toString().replaceAll('\n', r'\n')}';
  return text.length > width ? '${text.substring(0, width - 1)}…' : text;
}

/// Dumps the decoded grid so the parsed output can be eyeballed in the test
/// log. Run with `dart test -r expanded` to see it.
void logTable(SpreadsheetTable table) {
  print('sheet "$sheetName": ${table.maxRows} rows x ${table.maxCols} cols');
  for (var i = 0; i < table.rows.length; i++) {
    var cells = table.rows[i].map(_cell).join(' | ');
    print('[${i.toString().padLeft(2)}] $cells');
  }
}

void main() {
  group('Teaching plan sample excel file:', () {
    test('decode xlsx without update', () {
      var decoder = decode(filename, update: false);
      expect(decoder, isNotNull);
      expect(decoder.tables.isNotEmpty, isTrue);
      expect(decoder.tables.containsKey(sheetName), isTrue);

      var table = decoder.tables[sheetName]!;
      expect(table.maxRows, greaterThan(0));
      expect(table.maxCols, greaterThan(0));
      expect(table.rows.length, equals(table.maxRows));
    });

    test('decode xlsx with update', () {
      var decoder = decode(filename, update: true);
      expect(decoder, isNotNull);
      expect(decoder.tables.containsKey(sheetName), isTrue);

      var table = decoder.tables[sheetName]!;
      expect(table.maxRows, greaterThan(0));
      expect(table.maxCols, greaterThan(0));

      // Test reading rows and specific cells
      for (var row in table.rows) {
        expect(row, isA<List>());
      }
    });

    test('logs parsed data', () {
      var table = decode(filename).tables[sheetName]!;
      logTable(table);

      // Guard the dump above against silently logging an empty grid.
      expect(table.rows, isNotEmpty);
      expect(table.rows.every((row) => row.length == table.maxCols), isTrue);
    });

    test('grid shape and header row', () {
      var table = decode(filename).tables[sheetName]!;

      expect(table.maxRows, equals(42));
      expect(table.maxCols, equals(8));
      expect(table.rows.length, equals(42));
      expect(table.rows.first, equals(headers));
    });

    test('column value types', () {
      var table = decode(filename).tables[sheetName]!;
      var body = table.rows.skip(1);

      for (var row in body) {
        // Topic is blank on the row covered by a vertical merge.
        expect(row[topicCol], anyOf(isNull, isA<String>()));
        expect(row[subTopicCol], isA<String>());
        expect(row[courseOutcomeCol], matches(RegExp(r'^CO\d$')));
        expect(row[learningOutcomeCol], equals('UO1'));
        expect(row[hoursCol], isA<int>());
        expect(row[minutesCol], isA<int>());
        // Never a raw serial: every planning date renders as text.
        expect(row[planningDateCol], isA<String>());
        // ExecutionDate is an empty styled cell (`<c r="H2" s="11"/>`).
        expect(row[executionDateCol], isNull);
      }

      expect(body.map((row) => row[topicCol]).where((t) => t == null).length,
          equals(1));
    });

    test('PlanningDate serials decode as dates, not numbers', () {
      var table = decode(filename).tables[sheetName]!;

      serialBackedPlanningDates.forEach((index, expected) {
        expect(table.rows[index][planningDateCol], equals(expected),
            reason: 'row $index (numFmtId 58 date serial)');
      });
    });

    test('PlanningDate text cells pass through unchanged', () {
      var table = decode(filename).tables[sheetName]!;

      textBackedPlanningDates.forEach((index, expected) {
        expect(table.rows[index][planningDateCol], equals(expected),
            reason: 'row $index (shared string)');
      });
    });

    test('every PlanningDate row is accounted for', () {
      var table = decode(filename).tables[sheetName]!;
      var covered = {
        ...serialBackedPlanningDates.keys,
        ...textBackedPlanningDates.keys,
      };

      expect(covered.length, equals(table.rows.length - 1));
      expect(covered, equals({for (var i = 1; i < table.rows.length; i++) i}));
    });

    test('dateFormat option is applied to PlanningDate serials', () {
      var table = decode(filename, dateFormat: 'dd/MM/yyyy').tables[sheetName]!;

      expect(table.rows[1][planningDateCol], equals('09/07/2026'));
      expect(table.rows[13][planningDateCol], equals('10/05/2026'));
      expect(table.rows[39][planningDateCol], equals('12/11/2026'));

      // Text cells are not reformatted — they were never date serials.
      expect(table.rows[4][planningDateCol], equals('16/9/2026'));
    });

    test('update round-trip preserves decoded PlanningDate values', () {
      var original = decode(filename, update: true);
      var reopened = SpreadsheetDecoder.decodeBytes(original.encode());

      var before = original.tables[sheetName]!;
      var after = reopened.tables[sheetName]!;

      expect(after.maxRows, equals(before.maxRows));
      for (var i = 0; i < before.rows.length; i++) {
        expect(after.rows[i][planningDateCol],
            equals(before.rows[i][planningDateCol]),
            reason: 'row $i');
      }
    });
  });
}
