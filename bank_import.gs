/**
 * @file bank_import.gs
 * @description Módulo para importar y procesar estados de cuenta bancarios,
 * ya sea desde archivos PDF con texto o desde CSV. Normaliza los datos
 * y los inserta en la hoja "Bancos".
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Muestra un selector de archivos de Drive para que el usuario elija
 * el PDF o CSV del estado de cuenta a importar.
 */
function showBankFilePicker() {
  // Esta función usará el File Picker de Google para seleccionar un archivo.
  // Una vez seleccionado, llamará a la función de procesamiento principal.
  SpreadsheetApp.getUi().alert('Selector de archivos de banco (en desarrollo).');
}

/**
 * Orquesta el proceso de importación de un archivo de banco.
 * @param {string} fileId El ID del archivo de Drive seleccionado por el usuario.
 */
function processBankStatement(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  let movements = [];

  if (blob.getContentType() === MimeType.PDF) {
    movements = parseBankPdf(blob);
  } else if (blob.getContentType() === MimeType.CSV || blob.getContentType() === 'text/csv') {
    movements = parseBankCsv(blob);
  } else {
    SpreadsheetApp.getUi().alert('Error', 'Formato de archivo no soportado. Por favor, elija un PDF o CSV.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (movements.length > 0) {
    updateBankSheet(movements);
  } else {
    SpreadsheetApp.getUi().alert('Aviso', 'No se pudieron extraer movimientos del archivo. Si es un PDF sin texto, por favor conviértalo a CSV y vuelva a intentarlo.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Intenta extraer texto de un PDF y parsearlo línea por línea usando RegEx.
 * @param {GoogleAppsScript.Base.Blob} pdfBlob El blob del archivo PDF.
 * @returns {Array<Array<any>> | null} Un array de filas de movimientos, o null si no hay texto.
 */
function parseBankPdf(pdfBlob) {
  let text;
  try {
    text = pdfBlob.getDataAsString('UTF-8');
  } catch (e) {
    Logger.log(`Error extrayendo texto del PDF: ${e.message}`);
    return null;
  }

  if (!text || text.trim().length === 0) {
    Logger.log('El PDF no contiene una capa de texto extraíble.');
    return null;
  }

  const lines = text.split('\n');
  const movements = [];
  const bank = detectBank(text) || 'GENERICO';
  const patterns = BANK_PATTERNS[bank];

  if (!patterns) {
    Logger.log(`No se encontraron patrones de RegEx para el banco: ${bank}`);
    return [];
  }

  let current_date = null;
  lines.forEach(line => {
    line = line.trim();
    const dateMatch = line.match(patterns.DATE_REGEX);
    if (dateMatch) {
      current_date = dateMatch[0];
    }

    // Lógica para BBVA donde el cargo y abono están en columnas
    // Asumimos que si hay dos montos, el primero es cargo y el segundo abono
    const amounts = line.match(/\$([\d,]+\.\d{2})/g) || [];
    if (amounts.length > 0 && current_date) {
        let concept = line.replace(current_date, '').replace(/\$([\d,]+\.\d{2})/g, '').trim();
        let charge = 0;
        let deposit = 0;

        // Simplificación: si hay 2 montos, el primero es cargo, el segundo abono.
        // Si hay 1 monto, necesitamos más contexto (no implementado aquí).
        if(amounts.length === 2) {
            charge = parseFloat(amounts[0].replace(/[$,]/g, ''));
            deposit = parseFloat(amounts[1].replace(/[$,]/g, ''));
        } else if (amounts.length === 1) {
            // No podemos saber si es cargo o abono con certeza. Asumimos cargo.
            charge = parseFloat(amounts[0].replace(/[$,]/g, ''));
        }

        if (charge > 0) {
            movements.push(createMovementRow(current_date, concept, charge, 'Cargo'));
        }
        if (deposit > 0) {
            movements.push(createMovementRow(current_date, concept, deposit, 'Abono'));
        }
    }
  });

  return movements;
}

/**
 * Función de ayuda para crear una fila de movimiento normalizada.
 */
function createMovementRow(date, concept, amount, nature) {
    return [
      new Date(date),
      concept,
      '', // Referencia
      amount,
      nature,
      '', // Cuenta Banco
      '', // Etiquetas
      ''  // Folio Póliza
    ];
}

/**
 * Orquesta la importación de un archivo de banco (PDF o CSV) desde la carpeta designada.
 */
function triggerBankImport() {
  const ui = SpreadsheetApp.getUi();
  const folderId = getConfigValue('ID Carpeta PDFs Banco en Drive');
  if (!folderId) {
    ui.alert('Configuración Requerida', 'Por favor, ingrese el ID de la carpeta de Google Drive para los archivos de banco en la hoja "CFG".', ui.ButtonSet.OK);
    return;
  }

  const response = ui.prompt('Importar Archivo de Banco', 'Escriba el nombre exacto del archivo (ej: "estado_cuenta.pdf" o "movimientos.csv") que se encuentra en su carpeta de bancos de Drive.', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK || !response.getResponseText()) {
    return;
  }

  const fileName = response.getResponseText().toLowerCase();

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);

    if (!files.hasNext()) {
      ui.alert('Error', `No se encontró el archivo "${fileName}" en la carpeta de bancos especificada.`, ui.ButtonSet.OK);
      return;
    }

    const file = files.next();
    const blob = file.getBlob();
    let movements = [];
    let archiveFolderName = 'Procesados_Otros';

    if (fileName.endsWith('.pdf')) {
      movements = parseBankPdf(blob);
      archiveFolderName = 'Procesados_PDF';
    } else if (fileName.endsWith('.csv')) {
      movements = parseBankCsv(blob);
      archiveFolderName = 'Procesados_CSV';
    } else {
      ui.alert('Error', 'Formato de archivo no soportado. Use PDF o CSV.', ui.ButtonSet.OK);
      return;
    }

    if (movements && movements.length > 0) {
      updateBankSheet(movements);

      let archiveFolder;
      const archiveFolders = folder.getFoldersByName(archiveFolderName);
      archiveFolder = archiveFolders.hasNext() ? archiveFolders.next() : folder.createFolder(archiveFolderName);
      file.moveTo(archiveFolder);

      ui.alert('Éxito', `Se importaron ${movements.length} movimientos y el archivo fue archivado.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Aviso', 'No se pudieron extraer movimientos del archivo. Si es un PDF, puede que sea una imagen sin texto. Intente con un archivo CSV.', ui.ButtonSet.OK);
    }

  } catch (e) {
    Logger.log(e);
    ui.alert('Error', `Ocurrió un error: ${e.message}`, ui.ButtonSet.OK);
  }
}


/**
 * Parsea un archivo CSV de movimientos bancarios.
 * Asume un formato con encabezados: Fecha, Concepto, Referencia, Cargo, Abono
 * @param {GoogleAppsScript.Base.Blob} csvBlob El blob del archivo CSV.
 * @returns {Array<Array<any>>} Un array de filas de movimientos normalizados.
 */
function parseBankCsv(csvBlob) {
  const text = csvBlob.getDataAsString('UTF-8');
  const records = Utilities.parseCsv(text);
  const headers = records.shift().map(h => h.toLowerCase()); // Normalizar encabezados

  const dateIndex = headers.indexOf('fecha');
  const conceptIndex = headers.indexOf('concepto');
  const refIndex = headers.indexOf('referencia');
  const chargeIndex = headers.indexOf('cargo'); // Debe
  const depositIndex = headers.indexOf('abono'); // Haber

  if (dateIndex === -1 || conceptIndex === -1 || (chargeIndex === -1 && depositIndex === -1)) {
    throw new Error("El archivo CSV debe contener al menos las columnas 'Fecha', 'Concepto' y 'Cargo' o 'Abono'.");
  }

  const movements = [];
  records.forEach(row => {
    const charge = parseFloat(row[chargeIndex] || 0);
    const deposit = parseFloat(row[depositIndex] || 0);

    if(charge === 0 && deposit === 0) return; // Ignorar filas sin monto

    const amount = charge > 0 ? charge : deposit;
    const nature = charge > 0 ? 'Cargo' : 'Abono';

    // ["Fecha", "Concepto", "Referencia", "Monto", "Naturaleza", "Cuenta Banco", "Etiquetas", "Folio Póliza"]
    const normalizedRow = [
      new Date(row[dateIndex]),
      row[conceptIndex],
      refIndex > -1 ? row[refIndex] : '',
      amount,
      nature,
      '', // Cuenta Banco (se puede definir después)
      '', // Etiquetas
      ''  // Folio Póliza
    ];
    movements.push(normalizedRow);
  });

  return movements;
}

/**
 * Escribe los movimientos extraídos en la hoja "Bancos".
 * @param {Array<Array<any>>} movements Las filas de movimientos a insertar.
 */
function updateBankSheet(movements) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.BANKS);
  const startRow = sheet.getLastRow() + 1;

  sheet.getRange(startRow, 1, movements.length, movements[0].length).setValues(movements);
  SpreadsheetApp.getUi().alert(`Se importaron ${movements.length} movimientos a la hoja "Bancos".`);
}

/**
 * Detecta el banco basado en el contenido del texto.
 * @param {string} text Contenido del estado de cuenta.
 * @returns {string|null} El nombre del banco (ej. 'BBVA') o null.
 */
function detectBank(text) {
  if (/BBVA/i.test(text)) return 'BBVA';
  if (/Banorte/i.test(text)) return 'BANORTE';
  if (/Santander/i.test(text)) return 'SANTANDER';
  if (/Nu/i.test(text)) return 'NU';
  return null;
}
