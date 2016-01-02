# Spreadsheet
 
Spreadsheet is a library for reading spreadsheets as ODS and XLSX files.

## Usage

A simple usage example:

    import 'dart:io';
    import 'package:spreadsheet/spreadsheet.dart';

    main() {
      var bytes = new File.fromUri(fullUri).readAsBytesSync();
      var spreadsheetDecoder = new SpreadsheetDecoder.decodeBytes(bytes);
    }

## Features and bugs
Please file feature requests and bugs at the [issue tracker](tracker).

## License
Please see the [license file](LICENSE).
