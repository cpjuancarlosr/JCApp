/**
 * @OnlyCurrentDoc
 *
 * The above comment directs App Script to limit the scope of script authorization
 * to only the current document.
 */

/**
 * Creates a custom menu in the spreadsheet UI.
 * This function is a simple trigger that runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('SAT Parser')
      .addItem('Cargar XML...', 'showXmlImportDialog')
      .addSeparator()
      .addItem('Cargar PDF... (Próximamente)', 'pdfPlaceholderFunction')
      // .setDisabled(true) // .setDisabled is not a function on MenuItem, handled by placeholder
      .addToUi();
}

/**
 * Displays an HTML Service dialog in Google Sheets that allows users to upload files.
 */
function showXmlImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index')
      .setWidth(400)
      .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cargar Archivos XML');
}

/**
 * A placeholder function for the disabled PDF menu item.
 * It just shows an alert to the user.
 */
function pdfPlaceholderFunction() {
  SpreadsheetApp.getUi().alert('Función no implementada', 'La capacidad de procesar archivos PDF se añadirá en una futura actualización.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Processes the uploaded files, parses them, and writes the data to the active sheet.
 * This function is called from the client-side JavaScript from the dialog.
 *
 * @param {Object} formObject The form object from the client side, containing file contents.
 * @returns {Object} A status object with a message for the user.
 */
function processUploadedFiles(formObject) {
  try {
    const filesContent = formObject.files;
    if (!filesContent || filesContent.length === 0) {
      throw new Error("No files were uploaded.");
    }

    let allRows = [];
    filesContent.forEach(xmlContent => {
      const parsedRows = parseCfdiXml(xmlContent);
      allRows = allRows.concat(parsedRows);
    });

    if (allRows.length === 0) {
      return { status: 'warning', message: 'Could not extract any data from the files.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();

    // If the sheet is empty, add the header row first.
    if (lastRow === 0) {
      const header = getHeaderRow();
      allRows.unshift(header);
    }

    // Write all data to the sheet at once.
    const startRow = lastRow + 1;
    const range = sheet.getRange(startRow, 1, allRows.length, allRows[0].length);
    range.setValues(allRows);

    return { status: 'success', message: `Successfully processed ${filesContent.length} file(s) and added ${allRows.length} row(s).` };

  } catch (e) {
    Logger.log(e);
    // It's good practice to return a descriptive error to the client.
    return { status: 'error', message: 'An error occurred: ' + e.message };
  }
}
