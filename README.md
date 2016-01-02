# Spreadsheet
 
Spreadsheet is a library in dart language for reading spreadsheets for ODS and XLSX files.

## Usage

A simple usage example:

    import 'dart:io';
    import 'package:spreadsheet/spreadsheet.dart';

    main() {
      var bytes = new File.fromUri(fullUri).readAsBytesSync();
      var spreadsheetDecoder = new SpreadsheetDecoder.decodeBytes(bytes);
    }

## Not yet supported
This implementation doesn't support following features:
- annotations
- spanned rows
- spanned columns
- hidden rows (visible in resulting table)
- hidden columns (visible in resulting table)

## License
Please see the [license file](LICENSE).
