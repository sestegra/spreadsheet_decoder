import 'dart:io';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

main(List<String> args) {
  var bytes = new File(args[0]).readAsBytesSync();
  var decoder = new SpreadsheetDecoder.decodeBytes(bytes);
  for (var table in decoder.tables.keys) {
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }
}
