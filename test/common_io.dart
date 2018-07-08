import 'dart:convert';
import 'dart:io';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

List<int> _readBytes(String filename) {
  var fullUri = new Uri.file('test/files/$filename');
  return new File.fromUri(fullUri).readAsBytesSync();
}

String readBase64(String filename) {
  return base64Encode(_readBytes(filename));
}

SpreadsheetDecoder decode(String filename, {bool update = false}) {
  return new SpreadsheetDecoder.decodeBytes(_readBytes(filename), update: update, verify: true);
}

void save(String file, List<int> data) {
  new File(file)
    ..createSync(recursive: true)
    ..writeAsBytesSync(data);
}
