part of spreadsheet;

/// Read and parse XSLX spreadsheet
class XlsxDecoder extends SpreadsheetDecoder {
  /// Returns true if the Excel file is using the 1904 date epoch instead of the 1900 epoch.
  /// The Windows version of Excel generally uses the 1900 epoch while the Mac version of Excel generally uses the 1904 epoch.
  bool isUsing1904Date;

  XlsxDecoder(Archive archive) {
    this._archive = archive;
  }
}
