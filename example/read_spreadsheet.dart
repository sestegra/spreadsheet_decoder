import 'dart:io';
import 'package:path/path.dart';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

main(List<String> args) {
  var file = "test/files/test.ods";
  var bytes = new File(file).readAsBytesSync();
  var decoder = new SpreadsheetDecoder.decodeBytes(bytes, update: true);
  for (var table in decoder.tables.keys) {
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }

  var sheet = decoder.tables.keys.first;
  decoder
    ..updateCell(sheet, 0, 0, "New A")
    ..updateCell(sheet, 1, 0, "New B")
    ..updateCell(sheet, 2, 0, "New C")
    ..updateCell(sheet, 1, 1, "New D");

  new File(join("test/out/${basename(file)}"))
    ..createSync(recursive: true)
    ..writeAsBytesSync(decoder.encode());
}
