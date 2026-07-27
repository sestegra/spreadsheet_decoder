# Spreadsheet Decoder

[![Build Status](https://travis-ci.org/sestegra/spreadsheet_decoder.svg)](https://travis-ci.org/sestegra/spreadsheet_decoder?branch=master)
[![Coverage Status](https://coveralls.io/repos/sestegra/spreadsheet_decoder/badge.svg?branch=master)](https://coveralls.io/r/sestegra/spreadsheet_decoder?branch=master)
[![Pub version](https://img.shields.io/pub/v/spreadsheet_decoder.svg)](https://pub.dartlang.org/packages/spreadsheet_decoder)

Spreadsheet Decoder is a library for decoding and updating spreadsheets for ODS and XLSX files.

## Usage

### On server-side

    import 'dart:io';
    import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

    main() {
      var bytes = File.fromUri(fullUri).readAsBytesSync();
      var decoder = SpreadsheetDecoder.decodeBytes(bytes);
      var table = decoder.tables['Sheet1'];
      var values = table.rows[0];
      ...
      decoder.updateCell('Sheet1', 0, 0, 1337);
      File(join(fullUri).writeAsBytesSync(decoder.encode());
      ...
    }

### On client-side

    import 'dart:html';
    import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

    main() {
      var reader = FileReader();
      reader.onLoadEnd.listen((event) {
        var decoder = SpreadsheetDecoder.decodeBytes(reader.result);
        var table = decoder.tables['Sheet1'];
        var values = table.rows[0];
        ...
        decoder.updateCell('Sheet1', 0, 0, 1337);
        var bytes = decoder.encode();
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

For XLSX format, this implementation supports the native Excel formats for date, time and boolean type conversion, plus custom `<numFmt>` format codes declared in the workbook (e.g. `dd/mm/yyyy`).

Important: Excel often stores date cells as numeric serial values and only formats them for display. The decoder applies the workbook's format code to render the date as a string. If you need a fixed text representation independent of the workbook formatting, the source cell must be stored as text in the spreadsheet.

### Customising the date output format

By default, date cells (XLSX built-in numFmtIds 14-17, 22 and ODS `date` cells) are rendered as `yyyy-MM-dd` (e.g. `2008-07-21`). You can override this from the call site by passing `dateFormat`:

    var decoder = SpreadsheetDecoder.decodeBytes(
      bytes,
      dateFormat: 'dd/MM/yyyy', // → 21/07/2008
    );

Supported tokens (case-insensitive): `yyyy`, `yy`, `mm`/`m` (month — context-sensitive vs minutes), `dd`/`d`/`ddd`/`dddd`, `hh`/`h`, `ss`/`s`, `AM/PM`. Any other characters in the pattern (`-`, `/`, `.`, spaces, etc.) are emitted as literal separators.

For cells that use a custom `<numFmt>` format code declared in the workbook, the workbook's own format code is used unless you explicitly pass a non-default `dateFormat`, in which case your format wins.

## License

The MIT License, see [LICENSE](https://github.com/sestegra/spreadsheet_decoder/raw/master/LICENSE).
