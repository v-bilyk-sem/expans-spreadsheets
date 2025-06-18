/**
 * Adds a custom menu item so you can run the import from the sheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('EXPANS Functions')
    .addItem('Import all CSVs', 'importCSVsFromFolder')
    .addItem('Freeze 1st Row All Sheets', 'freezeFirstRowsAllSheets')
    .addItem('Crop Sheets to Data', 'cropAllSheetsToData')
    .addItem('Combine Non-Hidden Sheets', 'combineNonHiddenSheets')
    .addToUi();
}

/**
 * combineNonHiddenSheets:
 * - skips hidden sheets
 * - takes the first sheet’s header row as the master header
 * - for every sheet (including the first), grabs rows 2..last and prefixes with sheet name
 * - writes everything into a “Combined” sheet (re-creating it each time)
 */
function combineNonHiddenSheets() {
  const ss = SpreadsheetApp.getActive();
  const all = ss.getSheets().filter(s => !s.isSheetHidden());
  if (all.length === 0) {
    SpreadsheetApp.getUi().alert('No visible sheets to combine.');
    return;
  }
  // delete old Combined sheet if it exists
  const combinedName = 'Combined';
  let combined = ss.getSheetByName(combinedName);
  if (combined) ss.deleteSheet(combined);
  combined = ss.insertSheet(combinedName);

  // build header from first sheet
  const first = all[0];
  const lastCol = first.getLastColumn();
  const headerRow = first.getRange(1, 1, 1, lastCol).getValues()[0];
  combined
    .getRange(1, 1, 1, headerRow.length + 1)
    .setValues([['Sheet Name', ...headerRow]]);

  // gather data rows from every sheet (skipping row 1)
  const out = [];
  all.forEach(sheet => {
    const name = sheet.getName();
    const lr = sheet.getLastRow();
    if (lr < 2) return;            // nothing beyond header
    const data = sheet
      .getRange(2, 1, lr - 1, lastCol)
      .getValues();
    data.forEach(r => out.push([name, ...r]));
  });

  // write them out
  if (out.length) {
    combined
      .getRange(2, 1, out.length, headerRow.length + 1)
      .setValues(out);
  }
  SpreadsheetApp.getUi().alert('Sheets combined into “Combined”.');
}



/**
 * cropAllSheetsToData:
 * For each sheet, finds the last row/column with content,
 * then deletes all rows below and columns to the right.
 */
function cropAllSheetsToData() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();

    // Delete blank rows below data
    if (maxRows > lastRow) {
      sheet.deleteRows(lastRow + 1, maxRows - lastRow);
    }

    // Delete blank columns to the right of data
    if (maxCols > lastCol) {
      sheet.deleteColumns(lastCol + 1, maxCols - lastCol);
    }
  });
  SpreadsheetApp.getUi().alert('All sheets cropped to their data ranges.');
}

/**
 * Loops through every sheet in the spreadsheet
 * and sets the number of frozen rows to 1.
 */
function freezeFirstRowsAllSheets() {
  const ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(sheet => {
    sheet.setFrozenRows(1);
  });
  SpreadsheetApp.getUi().alert('First row frozen on all sheets!');
}

/**
 * Imports all CSV files in the specified Drive folder,
 * decoding as UTF-16LE and splitting on tabs.
 */
function importCSVsFromFolder() {
  // ▶︎ Change these to match your files:
  var FOLDER_ID = '1es643Qq-KpHfpOYbiSuBzpihYExLhZWs';
  var ENCODING  = 'UTF-16LE';  // for null-padded (UTF-16LE) CSVs
  var DELIMITER = '\t';        // tab-separated

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var files  = folder.getFilesByType(MimeType.CSV);

  while (files.hasNext()) {
    var file      = files.next();
    var fileName  = file.getName();
    var sheetName = fileName.replace(/\.csv$/i, '');

    // get or create the sheet
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clearContents();
    }

    // read & normalize the CSV text
    var blob      = file.getBlob();
    var csvString = blob.getDataAsString(ENCODING);

    // strip BOM if present
    if (csvString.charCodeAt(0) === 0xFEFF) {
      csvString = csvString.slice(1);
    }
    // remove any stray NULLs
    csvString = csvString.replace(/\u0000/g, '');

    // parse & write
    var data = Utilities.parseCsv(csvString, DELIMITER);
    if (data && data.length) {
      sheet
        .getRange(1, 1, data.length, data[0].length)
        .setValues(data);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Imported CSVs with ' + ENCODING +
    ' and tab delimiter.'
  );
}
