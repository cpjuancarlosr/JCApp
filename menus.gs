/**
 * @file menus.gs
 * @description Módulo dedicado a la creación y gestión del menú personalizado y la UI principal.
 * Centraliza la definición de la interfaz de usuario para facilitar su mantenimiento.
 *
 * @author Jules
 * @version 1.1.0
 */

/**
 * Se ejecuta cuando el usuario abre la hoja de cálculo.
 * Crea el menú personalizado y muestra la barra lateral principal.
 */
function onOpen() {
  createCustomMenu();
  showSidebar();
}

/**
 * Muestra la barra lateral principal de la aplicación.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar.html')
      .setTitle('JC Contable :: Panel de Control');
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Crea el menú "JC Contable" y lo agrega a la interfaz de usuario de la hoja de cálculo.
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('JC Contable');

  menu.addItem('▶ Mostrar Panel de Control', 'showSidebar');
  menu.addSeparator();
  menu.addItem('Configurar Hojas de Trabajo', 'setupSpreadsheet');

  // El resto del menú puede ser re-evaluado, ya que el panel de control es ahora la UI principal
  // Por ahora se dejan como placeholders

  const satMenu = ui.createMenu('SAT (CFDI)');
  satMenu.addItem('Configurar FIEL', 'placeholder_configureFiel');
  menu.addSubMenu(satMenu);

  const reportsMenu = ui.createMenu('Reportes Manuales');
  reportsMenu.addItem('Recalcular Estados Financieros', 'placeholder_recalculateReports');
  menu.addSubMenu(reportsMenu);

  menu.addSeparator();
  menu.addItem('Ayuda y Guía Rápida', 'placeholder_showHelp');

  menu.addToUi();
}

// --- Funciones Placeholder para mantener el menú funcional ---

function placeholder_configureFiel() {
  SpreadsheetApp.getUi().alert('Esta configuración se moverá a una sección dedicada en el futuro.');
}
function placeholder_recalculateReports() {
    SpreadsheetApp.getUi().alert('Use los botones del Panel de Control para generar reportes.');
}
function placeholder_showHelp() {
    SpreadsheetApp.getUi().alert('Guía rápida en desarrollo.');
}
