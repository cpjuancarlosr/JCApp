/**
 * @file menus.gs
 * @description Módulo dedicado a la creación y gestión del menú personalizado de la aplicación.
 * Centraliza la definición de la interfaz de usuario para facilitar su mantenimiento.
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Crea el menú "JC Contable" y lo agrega a la interfaz de usuario de la hoja de cálculo.
 * Esta función es llamada por onOpen() en Code.gs.
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('JC Contable');

  menu.addItem('▶ Configurar Hojas de Trabajo', 'setupSpreadsheet');

  const satMenu = ui.createMenu('SAT (CFDI)');
  satMenu.addItem('Configurar FIEL', 'placeholder_configureFiel');
  satMenu.addItem('Descargar CFDI de Mes/Año', 'placeholder_downloadCfdi');
  satMenu.addSeparator();
  satMenu.addItem('Procesar CFDI desde Carpeta', 'processCfdisFromDriveFolder');
  menu.addSubMenu(satMenu);

  const bankMenu = ui.createMenu('Bancos');
  bankMenu.addItem('Subir PDF/CSV Bancario', 'placeholder_uploadBankPdfs');
  menu.addSubMenu(bankMenu);

  menu.addSeparator();

  const accountingMenu = ui.createMenu('Procesos Contables');
  accountingMenu.addItem('Armar Pólizas', 'placeholder_buildPolicies');
  accountingMenu.addItem('Conciliar XML vs Bancos', 'placeholder_reconcile');
  accountingMenu.addItem('Recalcular Estados Financieros', 'placeholder_recalculateReports');
  menu.addSubMenu(accountingMenu);

  menu.addSeparator();

  menu.addItem('Exportar Reporte a PDF', 'placeholder_exportPdf');
  menu.addItem('Ayuda y Guía Rápida', 'placeholder_showHelp');

  menu.addToUi();
}

/**
 * Placeholder for a setup submenu.
 */
function setupSubMenu() {
  SpreadsheetApp.getUi().alert('Submenú de configuración.');
}
