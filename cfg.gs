/**
 * @file cfg.gs
 * @description Módulo de configuración. Centraliza constantes, nombres de hojas,
 * rangos con nombre y otros parámetros para que el resto del script no dependa
 * de valores hardcodeados.
 *
 * @author Jules
 * @version 1.0.0
 */

// Nombres de las hojas de cálculo
const SHEETS = {
  HOME: 'Inicio',
  CONFIG: 'CFG',
  ACCOUNTS: 'Cuentas',
  ENTITIES: 'Entidades',
  XML: 'XML',
  BANKS: 'Bancos',
  POLIZAS: 'Polizas'
};

// Nombres de rangos (Named Ranges)
// Se definirán programáticamente durante la configuración inicial.
const NAMED_RANGES = {
  CFG_XML_FOLDER_ID: 'CfgXmlFolderId',
  CFG_PDF_FOLDER_ID: 'CfgPdfFolderId',
  CFG_BANK_FOLDER_ID: 'CfgBankFolderId',
  CFG_CERT_PATH: 'CfgCertPath',
  CFG_KEY_PATH: 'CfgKeyPath',
  CFG_FIEL_PASSWORD: 'CfgFielPassword',
  // ... más rangos según sea necesario
};

// Mapeos y configuraciones específicas
// Ejemplo de mapeo para extracción de datos bancarios.
// NOTA: Estos patrones son ejemplos y pueden requerir ajustes.
const BANK_PATTERNS = {
  BBVA: {
    // Formato DD MMM. Ejemplo: 29 ENE
    DATE_REGEX: /^(\d{2}\s(?:ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC))/i,
    // Busca una línea que NO empiece con fecha y que tenga montos al final.
    CONCEPT_REGEX: /^(?!\d{2}\s(?:ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)).+?(?=\s\$|,)/i,
    // Busca un monto en formato $xx,xxx.xx al final de la línea.
    CHARGE_REGEX: /\$([\d,]+\.\d{2})\s*$/,
    DEPOSIT_REGEX: /\$([\d,]+\.\d{2})\s*$/
    // En BBVA, los cargos y abonos suelen estar en columnas separadas.
    // Una lógica más avanzada podría verificar la posición en la línea.
  },
  GENERICO: {
    // Un patrón genérico que podría funcionar en algunos casos.
    DATE_REGEX: /(\d{2}[\/-]\d{2}[\/-]\d{4})/i,
    CONCEPT_REGEX: /compra|pago|transferencia|deposito/i,
    CHARGE_REGEX: null, // No es posible un regex genérico fiable
    DEPOSIT_REGEX: null
  }
};

/**
 * Obtiene un valor de la hoja de configuración buscando el nombre del parámetro en la columna A.
 * @param {string} paramName El nombre del parámetro a buscar (ej. 'ID Carpeta XML en Drive').
 * @returns {any|null} El valor de la celda adyacente (columna B) o null si no se encuentra.
 */
function getConfigValue(paramName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (!configSheet) {
    throw new Error(`La hoja de configuración "${SHEETS.CONFIG}" no existe.`);
  }

  const data = configSheet.getRange('A1:B' + configSheet.getLastRow()).getValues();
  const param = data.find(row => row[0] === paramName);

  if (param) {
    return param[1]; // Devuelve el valor en la columna B
  }

  Logger.log(`Parámetro de configuración no encontrado: ${paramName}`);
  return null;
}

/**
 * Establece un valor en la hoja de configuración buscando el nombre del parámetro en la columna A.
 * @param {string} paramName El nombre del parámetro a buscar.
 * @param {any} value El valor a establecer en la celda adyacente (columna B).
 */
function setConfigValue(paramName, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
    if (!configSheet) {
    throw new Error(`La hoja de configuración "${SHEETS.CONFIG}" no existe.`);
  }

  const data = configSheet.getRange('A1:A' + configSheet.getLastRow()).getValues();
  const rowIndex = data.findIndex(row => row[0] === paramName);

  if (rowIndex !== -1) {
    configSheet.getRange(rowIndex + 1, 2).setValue(value);
  } else {
    throw new Error(`Parámetro de configuración no encontrado: ${paramName}`);
  }
}
