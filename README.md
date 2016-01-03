# Spreadsheet Decoder

[![Build Status](https://travis-ci.org/sestegra/spreadsheet_decoder.svg)](https://travis-ci.org/sestegra/spreadsheet_decoder?branch=master)
[![Coverage Status](https://coveralls.io/repos/sestegra/spreadsheet_decoder/badge.svg?branch=master)](https://coveralls.io/r/sestegra/spreadsheet_decoder?branch=master)

Spreadsheet Decoder is a library for decoding spreadsheets for ODS and XLSX files.

## Usage

### On server-side

    import 'dart:io';
    import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

    main() {
      var bytes = new File.fromUri(fullUri).readAsBytesSync();
      var decoder = new SpreadsheetDecoder.decodeBytes(bytes);
      var table = decoder.tables['Sheet1'];
      var values = table.rows[0];
      ...
    }

### On client-side

    import 'dart:html';
    import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

    main() {
      var reader = new FileReader();
      reader.onLoadEnd.listen((event) {
        var decoder = new SpreadsheetDecoder.decodeBytes(reader.result);
        var table = decoder.tables['Sheet1'];
        var values = table.rows[0];
        ...
      });
    }

## Features not yet supported
This implementation doesn't support following features:
- annotations
- spanned rows
- spanned columns
- hidden rows (visible in resulting tables)
- hidden columns (visible in resulting tables)
