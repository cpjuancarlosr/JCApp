/**
 * @file utils.gs
 * @description Módulo de utilidades. Contiene funciones auxiliares de propósito general
 * que son utilizadas por otros módulos en el proyecto.
 *
 * @author Jules
 * @version 1.1.0
 */

// Cache para almacenar la configuración y evitar lecturas repetidas de la hoja.
const SCRIPT_CACHE = CacheService.getScriptCache();
const CONFIG_CACHE_KEY = 'contabilidad_config';

/**
 * Obtiene los parámetros de la hoja "Configuración" y los devuelve como un objeto.
 * Utiliza caché para mejorar el rendimiento en ejecuciones sucesivas.
 *
 * @returns {object} Un objeto con la configuración, ej: { 'Periodo Contable Actual (YYYY-MM)': '2023-12', RFC: '...' }.
 */
function getContabilidadConfig() {
  const cachedConfig = SCRIPT_CACHE.get(CONFIG_CACHE_KEY);
  if (cachedConfig) {
    return JSON.parse(cachedConfig);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configuración');
  if (!sheet) {
    throw new Error('La hoja "Configuración" no se encuentra. Ejecute la configuración del menú.');
  }

  const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
  const config = {};

  data.forEach(row => {
    if (row[0] && row[1]) {
      config[row[0].trim()] = row[1].toString().trim();
    }
  });

  // Guardar en caché por 10 minutos.
  SCRIPT_CACHE.put(CONFIG_CACHE_KEY, JSON.stringify(config), 600);

  return config;
}

/**
 * Valida si un valor es un RFC válido (formato básico).
 * @param {string} rfc El RFC a validar.
 * @returns {boolean} True si el formato es válido, false en caso contrario.
 */
function isValidRfc(rfc) {
  if (!rfc || typeof rfc !== 'string') {
    return false;
  }
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
