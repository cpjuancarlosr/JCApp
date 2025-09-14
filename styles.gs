/**
 * @file styles.gs
 * @description Módulo para gestionar todos los estilos visuales y de formato de la hoja de cálculo.
 * Define la paleta de colores, fuentes y formatos numéricos para mantener una apariencia consistente.
 *
 * @author Jules
 * @version 1.0.0
 */

const THEME_COLORS = {
  ORANGE_ACCENT: '#DD9F86',
  BLUE_ACCENT: '#A6C2DB',
  TEXT: '#111111',
  BACKGROUND: '#FFFFFF',
  GRAY_LIGHT: '#F3F3F3' // Para cebras en tablas
};

const FONT_FAMILY = 'Rubik';

/**
 * Aplica el tema visual completo a toda la hoja de cálculo.
 * Establece la fuente predeterminada y otros estilos base.
 * Esta función es llamada por onOpen().
 */
function applyGlobalStyles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  allSheets.forEach(sheet => {
    // Establece la fuente para un rango de trabajo razonable para evitar exceder los límites de celdas.
    // 2000 filas y 52 columnas (hasta AZ) es un rango amplio y seguro.
    const rows = Math.min(sheet.getMaxRows(), 2000);
    const cols = Math.min(sheet.getMaxColumns(), 52);
    const range = sheet.getRange(1, 1, rows, cols);
    range.setFontFamily(FONT_FAMILY);
  });

  Logger.log(`Fuente '${FONT_FAMILY}' aplicada a todas las hojas.`);
}

/**
 * Formatea un rango para que parezca un título.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango a formatear.
 */
function formatAsTitle(range) {
  range
    .setFontFamily(FONT_FAMILY)
    .setFontSize(14)
    .setFontWeight('bold')
    .setFontColor(THEME_COLORS.TEXT)
    .setBackground(THEME_COLORS.BLUE_ACCENT);
}

/**
 * Formatea un rango para que parezca un encabezado de tabla.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango a formatear.
 */
function formatAsHeader(range) {
  range
    .setFontFamily(FONT_FAMILY)
    .setFontSize(11)
    .setFontWeight('bold')
    .setFontColor(THEME_COLORS.TEXT)
    .setBackground(THEME_COLORS.ORANGE_ACCENT)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
}

/**
 * Aplica formato de cebra a una tabla.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango de datos de la tabla.
 */
function applyZebraStriping(range) {
  range
    .setFontFamily(FONT_FAMILY)
    .setFontSize(10)
    .setFontColor(THEME_COLORS.TEXT)
    .setBackground(THEME_COLORS.BACKGROUND)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
    .getBanding()
    .setFirstRowColor(THEME_COLORS.ORANGE_ACCENT)
    .setSecondRowColor(THEME_COLORS.BACKGROUND)
    .setHeaderRowColor(THEME_COLORS.BLUE_ACCENT);
}

/**
 * Formato para números en formato de moneda MXN.
 * @param {GoogleAppsScript.Spreadsheet.Range} range El rango a formatear.
 */
function formatAsCurrency(range) {
  range.setNumberFormat('$#,##0.00');
}
