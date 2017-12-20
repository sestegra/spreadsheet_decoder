import 'dart:io';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

void main(List<String> args) {
  var path = args[0];
  var sheet = args.length > 1 ? args[1] : null;
  var data = new File(path).readAsBytesSync();
  var decoder = new SpreadsheetDecoder.decodeBytes(data, update: true);
  print(decoder.dumpXmlContent(sheet));
}
