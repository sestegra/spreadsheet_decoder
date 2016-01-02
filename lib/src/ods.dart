part of spreadsheet;

/// Read and parse ODS spreadsheet
class OdsDecoder extends SpreadsheetDecoder {
  OdsDecoder(Archive archive) {
    this._archive = archive;
  }
}
