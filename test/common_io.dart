import 'dart:convert';
import 'dart:io';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

List<int> _readBytes(String filename) {
  var fullUri = Uri.file('test/files/$filename');
  return File.fromUri(fullUri).readAsBytesSync();
}

String readBase64(String filename) {
  return base64Encode(_readBytes(filename));
}

SpreadsheetDecoder decode(String filename, {bool update = false}) {
  return SpreadsheetDecoder.decodeBytes(_readBytes(filename), update: update, verify: true);
}

void save(String file, List<int> data) {
  File(file)
    ..createSync(recursive: true)
    ..writeAsBytesSync(data);
}
