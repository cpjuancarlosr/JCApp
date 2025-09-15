/**
 * @OnlyCurrentDoc
 */

/**
 * Creates a custom menu in the spreadsheet UI.
 */
function onOpen() {
  setupConfigurationSheet();
  SpreadsheetApp.getUi()
      .createMenu('SAT Parser')
      .addItem('Cargar XML...', 'showXmlImportDialog')
      .addSeparator()
      .addItem('Cargar PDF... (Próximamente)', 'pdfPlaceholderFunction')
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
 * Placeholder for the disabled PDF menu item.
 */
function pdfPlaceholderFunction() {
  SpreadsheetApp.getUi().alert('Función no implementada', 'La capacidad de procesar archivos PDF se añadirá en una futura actualización.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Creates the 'Configuracion' sheet if it doesn't exist.
 */
function setupConfigurationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheetName = 'Configuracion';
  let sheet = ss.getSheetByName(configSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(configSheetName);
    sheet.getRange('A1').setValue('CONTACTOS').setFontWeight('bold').setBackground('#d9ead3');
    sheet.getRange('A2:C2').setValues([['RFC', 'Nombre', 'Tipo']]).setFontWeight('bold').setBackground('#d9ead3');
    sheet.getRange('E1').setValue('CATALOGO DE CUENTAS').setFontWeight('bold').setBackground('#d9ead3');
    sheet.getRange('E2:F2').setValues([['Codigo', 'Nombre de Cuenta']]).setFontWeight('bold').setBackground('#d9ead3');
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 300);
  }
}

/**
 * Processes uploaded files, updates the contacts database, and writes data to the sheet.
 */
function processUploadedFiles(formObject) {
  try {
    const filesContent = formObject.files;
    if (!filesContent || filesContent.length === 0) throw new Error("No files were uploaded.");

    let allSheetRows = [];
    let allContacts = [];
    filesContent.forEach(xmlContent => {
      const parsedData = parseCfdiXml(xmlContent);
      allSheetRows = allSheetRows.concat(parsedData.rows);
      allContacts = allContacts.concat(parsedData.contacts);
    });

    if (allSheetRows.length === 0) {
      return { status: 'warning', message: 'Could not extract any data from the files.' };
    }

    updateContactsDatabase(allContacts);

    const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Data');
    const lastRow = dataSheet.getLastRow();
    if (lastRow === 0) {
      const header = getHeaderRow();
      allSheetRows.unshift(header);
    }
    const range = dataSheet.getRange(lastRow + 1, 1, allSheetRows.length, allSheetRows[0].length);
    range.setValues(allSheetRows);

    return { status: 'success', message: `Successfully processed ${filesContent.length} file(s) and added ${allSheetRows.length} row(s).` };
  } catch (e) {
    Logger.log(e);
    return { status: 'error', message: 'An error occurred: ' + e.message };
  }
}

/**
 * Adds new contacts to the 'Configuracion' sheet if they don't already exist.
 */
function updateContactsDatabase(contacts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Configuracion');
  if (!configSheet) return;

  const contactRange = configSheet.getRange('A3:A' + configSheet.getMaxRows());
  const existingRfcs = new Set(contactRange.getValues().flat().filter(String));

  // Find the user's own RFC to avoid adding it to the contact list.
  // Heuristic: The user is the Emisor of Ingreso invoices.
  const userRfcCandidates = contacts
    .filter(c => c.tipoDeComprobante === 'I' && c.role === 'Emisor')
    .map(c => c.rfc);
  const userRfc = userRfcCandidates.length > 0 ? userRfcCandidates[0] : null;


  const newContacts = [];
  const uniqueNewRfcs = new Set();

  contacts.forEach(contact => {
    // Skip if the RFC is the user's own RFC, or if it already exists, or if it's already been added in this run.
    if (!contact.rfc || contact.rfc === userRfc || existingRfcs.has(contact.rfc) || uniqueNewRfcs.has(contact.rfc)) {
      return;
    }

    let tipo = '';
    if (contact.tipoDeComprobante === 'I' && contact.role === 'Receptor') {
      tipo = 'Cliente';
    } else if (contact.tipoDeComprobante === 'E' && contact.role === 'Emisor') {
      tipo = 'Proveedor';
    }

    if (tipo) {
      newContacts.push([contact.rfc, contact.nombre, tipo]);
      uniqueNewRfcs.add(contact.rfc);
    }
  });

  if (newContacts.length > 0) {
    const lastContactRow = configSheet.getRange('A:A').getValues().filter(String).length;
    configSheet.getRange(lastContactRow + 1, 1, newContacts.length, 3).setValues(newContacts);
  }
}
