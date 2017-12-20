import 'dart:io';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

SpreadsheetDecoder decode(String filename, {bool update = false}) {
  var fullUri = new Uri.file('test/files/$filename');
  return new SpreadsheetDecoder.decodeBytes(new File.fromUri(fullUri).readAsBytesSync(), update: update, verify: true);
}

void save(String file, List<int> data) {
  new File(file)
    ..createSync(recursive: true)
    ..writeAsBytesSync(data);
}
