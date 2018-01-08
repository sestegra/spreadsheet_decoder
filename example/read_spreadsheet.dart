import 'dart:io';
import 'package:path/path.dart';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

main(List<String> args) {
  var file = "test/files/test.xlsx";
  var bytes = new File(file).readAsBytesSync();
  var decoder = new SpreadsheetDecoder.decodeBytes(bytes, update: true);
  for (var table in decoder.tables.keys) {
    print(table);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }

  var sheet = decoder.tables.keys.first;
  decoder
    ..updateCell(sheet, 0, 0, 1337)
    ..updateCell(sheet, 1, 0, "New B")
    ..updateCell(sheet, 2, 0, "New C")
    ..updateCell(sheet, 1, 1, 42.3)
    ..insertRow(sheet, 1)
    ..insertRow(sheet, 13)
    ..updateCell(sheet, 0, 13, "A14")
    ..updateCell(sheet, 0, 12, "A13")
    ..insertColumn(sheet, 0)
    ..removeRow(sheet, 1)
    ..removeColumn(sheet, 2);

  new File(join("test/out/${basename(file)}"))
    ..createSync(recursive: true)
    ..writeAsBytesSync(decoder.encode());

  print("************************************************************");
  for (var table in decoder.tables.keys) {
    print(table);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }
}
