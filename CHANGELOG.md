## Unreleased
XLSX: decode dates with custom `<numFmt>` format codes (e.g. `dd/mm/yyyy`) using the workbook's own format code instead of leaking through as raw serial numbers. Removed the `raw` decoding mode in favour of always returning typed values.

Built-in date cells (numFmtId 14-17, 22) and ODS date cells now render as `yyyy-MM-dd` by default instead of an ISO timestamp like `2008-07-21T00:00:00.000`. Added a `dateFormat` parameter to `SpreadsheetDecoder.decodeBytes` and `SpreadsheetDecoder.decodeBuffer` so callers can choose any pattern (e.g. `dd/MM/yyyy`, `yyyy.MM.dd`); custom-numFmt date cells honour the caller's `dateFormat` when explicitly set, otherwise fall back to the workbook's own format code.

## 2.3.0
Update dependencies

## 2.2.0
Exclude phonetic information from cell value
Thanks to @okaxaki

## 2.1.1
Fix conflicting archive dependency for using integration_test package

## 2.1.0
Update dependencies

## 2.0.2
Fix breaking changes from archive 3.2.0

## 2.0.1
Fix for absolute path target

## 2.0.0
Release NNBD

## 2.0.0-nullsafety.0
Convert to NNBD

## 1.2.0
Add pedantic linter

## 1.1.1
Update dependencies

## 1.1.0
Fix insertion on empty sheet
Fix unnecessary string escape

## 1.0.1
Fix parsing on converted Numbers documents

## 1.0.0
Dart 2 support

## 0.4.3
Upgrade to xml 3.0.0

## 0.4.3-alpha
Add insert/remove columns/rows

## 0.4.2-alpha
Add extension getter

## 0.4.2-alpha
Add mediaType getter

## 0.4.0-alpha
Add updateCell method to update existing cells content

## 0.3.2
Fix string parsing on xlsx file
Fix sheet target selection

## 0.3.1
Fix misalignment of decoded value with empty columns

## 0.3.0
Output typed data (number, string, date, time, boolean)

## 0.2.0
Add XLSX decoder

## 0.1.0
Add ODS decoder

## 0.0.1
Initial release
