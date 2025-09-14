/**
 * @file utils.gs
 * @description Módulo de utilidades. Contiene funciones auxiliares de propósito general
 * que son utilizadas por otros módulos en el proyecto. Por ejemplo, validación de datos,
 * formato de fechas, creación de bitácoras, etc.
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Valida si un valor es un RFC válido (formato básico).
 * @param {string} rfc El RFC a validar.
 * @returns {boolean} True si el formato es válido, false en caso contrario.
 */
function isValidRfc(rfc) {
  if (!rfc || typeof rfc !== 'string') {
    return false;
  }
  // Expresión regular para RFC de persona física o moral (sin validación de homoclave).
  const rfcRegex = /^[A-Z&Ñ]{3,4}\d{6}[A-Z\d]{3}$/i;
  return rfcRegex.test(rfc);
}

/**
 * Formatea un objeto Date a un string 'YYYY-MM-DD'.
 * @param {Date} date El objeto Date a formatear.
 * @returns {string} La fecha formateada.
 */
function formatDate(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Escribe un mensaje en una hoja de bitácora (log).
 * @param {string} message El mensaje a registrar.
 */
function logMessage(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('Bitacora');

  if (!logSheet) {
    logSheet = ss.insertSheet('Bitacora');
    logSheet.appendRow(['Timestamp', 'Mensaje']);
  }

  const timestamp = new Date();
  logSheet.appendRow([timestamp, message]);
}

/**
 * Crea una validación de datos en una celda para que solo acepte valores de una lista.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango donde se aplicará la validación.
 * @param {Array<string>} allowedValues La lista de valores permitidos.
 * @param {string} helpText El mensaje de ayuda que se muestra al editar la celda.
 */
function createDropdownValidation(range, allowedValues, helpText) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowedValues)
    .setAllowInvalid(false)
    .setHelpText(helpText)
    .build();
  range.setDataValidation(rule);
}
