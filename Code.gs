/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a custom menu in the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('SAT Parser')
      .addItem('Cargar XML...', 'showXmlImportDialog')
      .addToUi();
}

/**
 * Displays the file upload dialog.
 */
function showXmlImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index').setWidth(400).setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cargar Archivos XML');
}

/**
 * Main function called from the dialog to process uploaded XML files.
 */
function processUploadedFiles(formObject) {
  try {
    const filesContent = formObject.files;
    if (!filesContent || filesContent.length === 0) throw new Error("No files were uploaded.");

    const incomeRows = [];
    const expenseRows = [];

    filesContent.forEach(xmlContent => {
      const parsedData = parseCfdiXml(xmlContent);
      if (parsedData.tipoDeComprobante === 'I') {
        incomeRows.push(parsedData.rowData);
      } else if (parsedData.tipoDeComprobante === 'E') {
        expenseRows.push(parsedData.rowData);
      }
    });

    if (incomeRows.length > 0) {
      writeDataToSheet('XML_I', incomeRows);
    }
    if (expenseRows.length > 0) {
      writeDataToSheet('XML_E', expenseRows);
    }

    const message = `Processed ${filesContent.length} files. Added ${incomeRows.length} income rows and ${expenseRows.length} expense rows.`;
    return { status: 'success', message: message };

  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'An error occurred: ' + e.message };
  }
}

/**
 * Writes an array of rows to a specified sheet, creating it if necessary.
 * Also triggers formula replication for the new rows.
 */
function writeDataToSheet(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  let startRow = sheet.getLastRow() + 1;
  if (startRow === 1) {
    sheet.appendRow(getNewHeaderRow());
    startRow++; // Data starts on row 2
  }

  const range = sheet.getRange(startRow, 1, rows.length, rows[0].length);
  range.setValues(rows);

  // Replicate formulas after writing data
  replicateFormulas(sheet, startRow, rows.length);
}

/**
 * Replicates formulas from the first data row (row 2) to newly added rows.
 * It copies formulas from column P (16) to the last column.
 */
function replicateFormulas(sheet, startRow, numRows) {
  const lastCol = sheet.getLastColumn();
  const formulaStartCol = 16; // Column P

  if (lastCol < formulaStartCol || sheet.getLastRow() < 2 || numRows === 0) {
    return; // No formulas to copy or no data rows exist yet.
  }

  // Get the formulas from the first data row (row 2)
  const formulaRange = sheet.getRange(2, formulaStartCol, 1, lastCol - formulaStartCol + 1);
  const formulas = formulaRange.getFormulasR1C1();

  // Get the target range for the new rows
  const targetRange = sheet.getRange(startRow, formulaStartCol, numRows, formulas[0].length);

  // Set the formulas. R1C1 notation handles relative references correctly.
  targetRange.setFormulasR1C1(formulas);
}
