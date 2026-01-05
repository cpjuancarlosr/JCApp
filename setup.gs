/**
 * @file setup.gs
 * @description Módulo para la configuración inicial de la hoja de cálculo.
 * Crea las hojas necesarias y establece los valores predeterminados.
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Función principal que orquesta la creación y configuración de todas las hojas de trabajo.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = {
    'Configuración': setupConfiguracionSheet,
    'Catálogo de Cuentas': setupCatalogoSheet,
    'Mapeo_ProdServ_Cuenta': setupMapeoSheet,
    'Pólizas': setupPolizasSheet,
    'Log_Procesados': setupLogSheet,
    'Balanza de Comprobación': setupBalanzaSheet,
    'Balance General': setupBalanceGeneralSheet,
    'Estado de Resultados': setupEstadoResultadosSheet,
    'Auxiliar de Cuenta': setupAuxiliarSheet
  };

  // Crear cada hoja si no existe y aplicar su configuración
  for (const sheetName in sheets) {
    if (!ss.getSheetByName(sheetName)) {
      const newSheet = ss.insertSheet(sheetName);
      sheets[sheetName](newSheet);
      Logger.log(`Hoja "${sheetName}" creada y configurada.`);
    } else {
      Logger.log(`Hoja "${sheetName}" ya existe. Se omite la creación.`);
    }
  }

  SpreadsheetApp.getUi().alert('La configuración y verificación de las hojas de trabajo ha finalizado.');
}

/**
 * Configura la hoja "Configuración".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupConfiguracionSheet(sheet) {
  sheet.getRange('A1:B1').setValues([['Parámetro', 'Valor']]).setFontWeight('bold');
  sheet.getRange('A2:B4').setValues([
    ['Periodo Contable Actual (YYYY-MM)', '2023-12'],
    ['RFC de la Empresa', 'INGRESA TU RFC AQUÍ'],
    ['Periodos Cerrados (lista separada por comas)', '2023-11,2023-10']
  ]);
  sheet.autoResizeColumns(1, 2);
}

/**
 * Configura la hoja "Catálogo de Cuentas".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupCatalogoSheet(sheet) {
  const headers = ['Cuenta', 'Nombre', 'Tipo', 'Subtipo', 'Naturaleza'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Configura la hoja "Mapeo_ProdServ_Cuenta".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupMapeoSheet(sheet) {
  const headers = ['ClaveProdServ', 'UsoCFDI', 'CuentaContable'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Configura la hoja "Pólizas".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupPolizasSheet(sheet) {
  const headers = ['Fecha', 'PolizaID', 'UUID_CFDI', 'Cuenta', 'Concepto', 'Debe', 'Haber'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Configura la hoja "Log_Procesados".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupLogSheet(sheet) {
  const headers = ['UUID', 'FechaProceso'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Configura la hoja "Balanza de Comprobación".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupBalanzaSheet(sheet) {
  const headers = ['Cuenta', 'Nombre', 'Saldo Inicial', 'Debe', 'Haber', 'Saldo Final'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/**
 * Configura la hoja "Balance General".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupBalanceGeneralSheet(sheet) {
  // Se deja vacía para ser llenada por el reporte.
  sheet.getRange('A1').setValue('Balance General').setFontWeight('bold');
}

/**
 * Configura la hoja "Estado de Resultados".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupEstadoResultadosSheet(sheet) {
    // Se deja vacía para ser llenada por el reporte.
  sheet.getRange('A1').setValue('Estado de Resultados').setFontWeight('bold');
}


/**
 * Configura la hoja "Auxiliar de Cuenta".
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja a configurar.
 */
function setupAuxiliarSheet(sheet) {
  const headers = ['Fecha', 'PolizaID', 'Concepto', 'Debe', 'Haber', 'Saldo'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.getRange('A2').setValue('Seleccione una cuenta y un periodo para generar el reporte.');
}
