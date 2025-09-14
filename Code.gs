/**
 * @file Code.gs
 * @description Punto de entrada principal de la aplicación. Gestiona los disparadores (triggers)
 * como onOpen() y onEdit() y contiene las funciones globales de alto nivel que se
 * asignan a los elementos del menú o botones.
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Se ejecuta cuando el libro de Google Sheets se abre.
 * Crea el menú personalizado y aplica estilos iniciales.
 */
function onOpen() {
  createCustomMenu();
  applyGlobalStyles();
}

/**
 * Se ejecuta cuando se cierra la hoja de cálculo.
 * Limpia propiedades de usuario sensibles como la contraseña de la FIEL.
 */
function onClose() {
  // Limpiar la contraseña de la FIEL de las propiedades del usuario
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('fielPassword');
  Logger.log('Propiedades de sesión (FIEL password) limpiadas.');
}

// Aquí se agregarán las funciones que son llamadas directamente por los menús
// y que a su vez orquestan llamadas a otros módulos.

function placeholder_configureFiel() {
  showFielConfigurationUi(); // Llama a la función en sat_connector.gs
}

function placeholder_downloadCfdi() {
  SpreadsheetApp.getUi().alert('Función "Descargar CFDI" no implementada aún.');
}

function placeholder_uploadBankPdfs() {
  triggerBankImport(); // Llama a la función unificada en bank_import.gs
}

function placeholder_buildPolicies() {
  generatePolicies(); // Llama a la función principal en polizas.gs
}

function placeholder_reconcile() {
  SpreadsheetApp.getUi().alert('Función "Conciliar" no implementada aún.');
}

function placeholder_recalculateReports() {
  setupFinancialReports(); // Llama a la función en reports.gs
}

function placeholder_exportPdf() {
  exportReportAsPdf(); // Llama a la función principal en pdf_export.gs
}

function placeholder_showHelp() {
  SpreadsheetApp.getUi().alert('Función "Ayuda" no implementada aún.');
}

/**
 * Crea y formatea todas las hojas de cálculo necesarias para la aplicación.
 * Se puede ejecutar desde el menú de configuración.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert('Confirmación de Configuración', 'Este proceso creará y formateará las hojas de trabajo necesarias. ¿Desea continuar?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    return;
  }

  const sheetsToCreate = {
    [SHEETS.HOME]: [],
    [SHEETS.CFG]: ['Parámetro', 'Valor'],
    [SHEETS.CUENTAS]: ['Código', 'Nombre Cuenta', 'Naturaleza (D/A)', 'Nivel', 'Agrupador SAT'],
    [SHEETS.ENTIDADES]: ['RFC', 'Razón Social', 'Tipo (Cliente/Proveedor)', 'Cuenta Contable', 'Clave Prod/Serv SAT', 'Unidad SAT', 'Tasa IVA Predeterminada'],
    [SHEETS.XML]: ['UUID', 'Emisor RFC', 'Receptor RFC', 'Tipo', 'Fecha', 'Método Pago', 'Forma Pago', 'Subtotal', 'Impuestos Trasladados', 'Total', 'Uso CFDI', 'Moneda', 'CFDI Relacionado'],
    [SHEETS.BANCOS]: ['Fecha', 'Concepto', 'Referencia', 'Monto', 'Naturaleza (Cargo/Abono)', 'Cuenta Banco', 'Etiquetas', 'Folio Póliza'],
    [SHEETS.POLIZAS]: ['Fecha Póliza', 'Tipo', 'Folio', 'Cuenta', 'Subcuenta', 'Concepto', 'Referencia', 'Debe', 'Haber', 'UUID Relacionado', 'Estatus']
  };

  for (const sheetName in sheetsToCreate) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log(`Hoja "${sheetName}" creada.`);
    }

    const headers = sheetsToCreate[sheetName];
    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      // Aplicar estilo de encabezado
      formatAsHeader(sheet.getRange(1, 1, 1, headers.length));
      sheet.setFrozenRows(1);
    }
  }

  // Ocultar hoja por defecto si existe
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    defaultSheet.hideSheet();
  }

  // Pre-llenar la hoja de configuración
  populateInitialConfig();

  applyGlobalStyles(); // Re-aplicar estilos
  ui.alert('Configuración completada', 'Se han creado y formateado las hojas de trabajo.', ui.ButtonSet.OK);
}

/**
 * Llena la hoja de CFG con los parámetros iniciales.
 */
function populateInitialConfig() {
  const cfgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CFG);
  cfgSheet.clear(); // Limpiar antes de poblar

  const params = [
    ['ID Carpeta XML en Drive', ''],
    ['ID Carpeta PDFs Banco en Drive', ''],
    ['Ruta Archivo .CER en Drive', ''],
    ['Ruta Archivo .KEY en Drive', ''],
    ['Contraseña FIEL', 'NO GUARDAR AQUÍ - USAR MENÚ'],
    ['Tasa IVA General (%)', '16'],
    ['Tasa IVA Frontera (%)', '8'],
    ['Activar Sincronización Automática', 'NO'],
  ];

  cfgSheet.getRange(1, 1, params.length, 2).setValues(params);
  formatAsHeader(cfgSheet.getRange("A1:B1"));
  // Proteger la celda de la contraseña como recordatorio visual
  const passCell = cfgSheet.getRange('B5');
  passCell.setBackground('#f4c7c3').setFontColor('#990000').protect().setWarningOnly(true);
}
