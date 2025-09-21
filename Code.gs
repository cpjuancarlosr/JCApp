/**
 * @OnlyCurrentDoc
 */

// =================================================================
// UI & SETUP FUNCTIONS
// =================================================================

/**
 * Creates the add-on menu when the spreadsheet is opened.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const satMenu = ui.createMenu('SAT Herramientas')
    .addItem('Descargar CFDI desde SAT...', 'runDescargaMasivaCfdi')
    .addSeparator()
    .addItem('Configurar Credenciales...', 'showFielSetupDialog');

  const manualMenu = ui.createMenu('Carga Manual')
    .addItem('Cargar XMLs (locales)...', 'showLocalXmlUploadDialog');
    // PDF Placeholder can be added back later if needed

  satMenu.addToUi();
  manualMenu.addToUi();
}

/**
 * Shows a dialog for setting the FIEL credentials.
 */
function showFielSetupDialog() {
  const html = HtmlService.createHtmlOutputFromFile('setup')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configurar Credenciales FIEL');
}

/**
 * Shows the dialog for uploading local XML files.
 */
function showLocalXmlUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Cargar Archivos XML Locales');
}


/**
 * Saves the FIEL credentials to script properties.
 */
function saveFielCredentials(form) {
  try {
    if (!form.rfc || !form.certificatePem || !form.privateKeyPem || !form.issuerName || !form.serialNumber) {
      throw new Error('Todos los campos son obligatorios.');
    }
    PropertiesService.getScriptProperties().setProperties({
      'FIEL_RFC': form.rfc,
      'FIEL_CERT_PEM': form.certificatePem,
      'FIEL_PRIVATE_KEY_PEM': form.privateKeyPem,
      'FIEL_ISSUER_NAME': form.issuerName,
      'FIEL_CERT_SERIAL': form.serialNumber
    });
    return { status: 'success', message: 'Credenciales guardadas correctamente.' };
  } catch (e) {
    return { status: 'error', message: 'Error: ' + e.message };
  }
}

// =================================================================
// GLOBAL CONSTANTS & CONFIG
// =================================================================

const SAT_NS = {
  SOAP: 'http://schemas.xmlsoap.org/soap/envelope/',
  DES: 'http://DescargaMasivaTerceros.sat.gob.mx',
  WSSE: 'http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd',
  WSU: 'http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd',
  DS: 'http://www.w3.org/2000/09/xmldsig#',
};

const SAT_ENDPOINTS = {
  PRODUCTION: { /* ... endpoints ... */ },
  TEST: { /* ... endpoints ... */ },
};
// The full SAT_ENDPOINTS object from user's code would be here. For brevity, it's omitted.
SAT_ENDPOINTS.PRODUCTION = {
    autenticacion: 'https://cfdidescargamasivasolicitud.clouda.sat.gob.mx/Autenticacion/Autenticacion.svc',
    solicita: 'https://cfdidescargamasivasolicitud.clouda.sat.gob.mx/SolicitaDescargaService.svc',
    verifica: 'https://cfdidescargamasivasolicitud.clouda.sat.gob.mx/VerificaSolicitudDescargaService.svc',
    descarga: 'https://cfdidescargamasiva.clouda.sat.gob.mx/DescargaMasivaService.svc',
    soapAction: {
      autentica: 'http://DescargaMasivaTerceros.gob.mx/IAutenticacion/Autentica',
      solicitaEmitidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaEmitidos',
      solicitaRecibidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaRecibidos',
      verifica: 'http://DescargaMasivaTerceros.sat.gob.mx/IVerificaSolicitudDescargaService/VerificaSolicitudDescarga',
      descarga: 'http://DescargaMasivaTerceros.sat.gob.mx/IDescargaMasivaTercerosService/Descargar',
    },
  };
SAT_ENDPOINTS.TEST = {
    autenticacion: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/Autenticacion/Autenticacion.svc',
    solicita: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/SolicitaDescargaService.svc',
    verifica: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/VerificaSolicitudDescargaService.svc',
    descarga: 'https://pruebassolicituddescargamasiva.clouda.sat.gob.mx/DescargaMasivaService.svc',
    soapAction: SAT_ENDPOINTS.PRODUCTION.soapAction
};


const DEFAULT_OPTIONS = { /* ... options ... */ };

function getScriptConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    certificatePem: props.getProperty('FIEL_CERT_PEM') || '',
    privateKeyPem: props.getProperty('FIEL_PRIVATE_KEY_PEM') || '',
    certificateIssuer: props.getProperty('FIEL_ISSUER_NAME') || '',
    certificateSerial: props.getProperty('FIEL_CERT_SERIAL') || '',
    rfc: props.getProperty('FIEL_RFC') || '',
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
    driveFolderId: props.getProperty('GOOGLE_DRIVE_FOLDER_ID') || '',
    environment: (props.getProperty('SAT_ENVIRONMENT') || 'PRODUCTION').toUpperCase(),
    pollIntervalSeconds: parseInt(props.getProperty('SAT_POLL_INTERVAL_SECONDS') || '60', 10),
    maxWaitMinutes: parseInt(props.getProperty('SAT_MAX_WAIT_MINUTES') || '30', 10),
  };
}

function assertConfig(config) { /* ... assertion logic ... */ }
// For brevity, the full function is omitted. It checks for missing properties.
assertConfig = function(config) {
  const missing = [];
  if (!config.certificatePem) missing.push('FIEL_CERT_PEM');
  if (!config.privateKeyPem) missing.push('FIEL_PRIVATE_KEY_PEM');
  if (!config.certificateIssuer) missing.push('FIEL_ISSUER_NAME');
  if (!config.certificateSerial) missing.push('FIEL_CERT_SERIAL');
  if (!config.rfc) missing.push('FIEL_RFC');
  if (!config.spreadsheetId) missing.push('GOOGLE_SHEETS_ID');
  if (missing.length) {
    throw new Error('Faltan propiedades de script obligatorias: ' + missing.join(', '));
  }
  if (!SAT_ENDPOINTS[config.environment]) {
    throw new Error('SAT_ENVIRONMENT debe ser PRODUCTION o TEST');
  }
}


// =================================================================
// UNIFIED PARSING & SHEET WRITING LOGIC
// =================================================================

/**
 * Defines the single, unified header row for all XML data.
 */
function getUnifiedHeaderRow() {
  return [
    'Periodo', 'Fecha', 'Serie', 'Folio', 'UUID', 'RFC_Emisor', 'RFC_Receptor',
    'Metodo(PUE/PPD)', 'TasaIVA(0/0.08/0.16)', 'Subtotal', 'Descuento',
    'IVA_Trasladado', 'IEPS', 'Ret_ISR', 'Ret_IVA', 'Total'
  ];
}

/**
 * Helper function to safely get an attribute value from an element.
 */
function getSafeAttribute(element, attributeName) {
  if (!element) return '';
  const attribute = element.getAttribute(attributeName);
  return attribute ? attribute.getValue() : '';
}

/**
 * The single, unified parser for any CFDI XML content.
 */
function parseAndFormatCfdi(xmlContent) {
  const document = XmlService.parse(xmlContent);
  const root = document.getRootElement();
  const cfdi = root.getNamespace();
  const tfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

  const comprobante = root;
  const emisor = comprobante.getChild('Emisor', cfdi);
  const receptor = comprobante.getChild('Receptor', cfdi);
  const timbre = comprobante.getChild('Complemento', cfdi).getChild('TimbreFiscalDigital', tfd);
  const impuestosNode = comprobante.getChild('Impuestos', cfdi);

  let ivaTrasladado = 0, ieps = 0, isrRetenido = 0, ivaRetenido = 0, tasaIVA = 0;
  const subtotal = parseFloat(getSafeAttribute(comprobante, 'SubTotal')) || 0;

  if (impuestosNode) {
    const traslados = impuestosNode.getChild('Traslados', cfdi);
    if (traslados) {
      traslados.getChildren('Traslado', cfdi).forEach(t => {
        const impuesto = getSafeAttribute(t, 'Impuesto');
        const importe = parseFloat(getSafeAttribute(t, 'Importe')) || 0;
        if (impuesto === '002') ivaTrasladado += importe;
        else if (impuesto === '003') ieps += importe;
      });
    }
    const retenciones = impuestosNode.getChild('Retenciones', cfdi);
    if (retenciones) {
      retenciones.getChildren('Retencion', cfdi).forEach(r => {
        const impuesto = getSafeAttribute(r, 'Impuesto');
        const importe = parseFloat(getSafeAttribute(r, 'Importe')) || 0;
        if (impuesto === '001') isrRetenido += importe;
        else if (impuesto === '002') ivaRetenido += importe;
      });
    }
  }

  if (subtotal > 0 && ivaTrasladado > 0) {
      const effectiveRate = (ivaTrasladado / subtotal);
      if (effectiveRate > 0.12) tasaIVA = 0.16;
      else if (effectiveRate > 0.04) tasaIVA = 0.08;
      else tasaIVA = 0;
  }

  const fecha = getSafeAttribute(comprobante, 'Fecha').split('T')[0];
  const period = fecha ? fecha.substring(0, 7) : '';
  const tipoDeComprobante = getSafeAttribute(comprobante, 'TipoDeComprobante');

  const rowData = [
    period, fecha, getSafeAttribute(comprobante, 'Serie'), getSafeAttribute(comprobante, 'Folio'),
    getSafeAttribute(timbre, 'UUID'), getSafeAttribute(emisor, 'Rfc'), getSafeAttribute(receptor, 'Rfc'),
    getSafeAttribute(comprobante, 'MetodoPago'), tasaIVA, subtotal.toFixed(2),
    getSafeAttribute(comprobante, 'Descuento') || '0.00', ivaTrasladado.toFixed(2), ieps.toFixed(2),
    isrRetenido.toFixed(2), ivaRetenido.toFixed(2), getSafeAttribute(comprobante, 'Total'),
  ];

  return { tipoDeComprobante: tipoDeComprobante, rowData: rowData };
}

/**
 * Writes an array of rows to a specified sheet and replicates formulas.
 */
function writeDataToSheet(sheetName, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  let startRow = sheet.getLastRow() + 1;
  if (startRow === 1) {
    sheet.appendRow(getUnifiedHeaderRow());
    startRow++;
  }

  if (rows.length > 0) {
    const range = sheet.getRange(startRow, 1, rows.length, rows[0].length);
    range.setValues(rows);
    replicateFormulas(sheet, startRow, rows.length);
  }
}

/**
 * Replicates formulas from row 2 to newly added rows.
 */
function replicateFormulas(sheet, startRow, numRows) {
  const lastCol = sheet.getLastColumn();
  const formulaStartCol = 16; // Column P

  if (lastCol < formulaStartCol || sheet.getLastRow() < 2 || numRows === 0) return;

  const formulaRange = sheet.getRange(2, formulaStartCol, 1, lastCol - formulaStartCol + 1);
  const formulas = formulaRange.getFormulasR1C1();

  const targetRange = sheet.getRange(startRow, formulaStartCol, numRows, formulas[0].length);
  targetRange.setFormulasR1C1(formulas);
}

// =================================================================
// LOCAL FILE UPLOAD WORKFLOW
// =================================================================

/**
 * Main function to process locally uploaded XML files.
 */
function processLocalFiles(formObject) {
  try {
    const filesContent = formObject.files;
    if (!filesContent || filesContent.length === 0) throw new Error("No files were uploaded.");

    const incomeRows = [];
    const expenseRows = [];

    filesContent.forEach(xmlContent => {
      const parsedData = parseAndFormatCfdi(xmlContent);
      if (parsedData.tipoDeComprobante === 'I') {
        incomeRows.push(parsedData.rowData);
      } else if (parsedData.tipoDeComprobante === 'E') {
        expenseRows.push(parsedData.rowData);
      }
    });

    if (incomeRows.length > 0) writeDataToSheet('XML_I', incomeRows);
    if (expenseRows.length > 0) writeDataToSheet('XML_E', expenseRows);

    const message = `Carga manual procesada. Se agregaron ${incomeRows.length} filas de ingresos y ${expenseRows.length} filas de gastos.`;
    return { status: 'success', message: message };

  } catch (e) {
    Logger.log('Error en processLocalFiles: ' + e.toString());
    return { status: 'error', message: 'Error: ' + e.message };
  }
}


// =================================================================
// SAT DOWNLOAD WORKFLOW (Refactored)
// =================================================================

function runDescargaMasivaCfdi() { /* ... unchanged ... */ }
function descargarCfdiMasivo(userOptions) { /* ... main logic ... */ }
// For brevity, the full functions are omitted.
// The key change is inside `parsearPaquete` and the end of `descargarCfdiMasivo`.
runDescargaMasivaCfdi = function() {
  const config = getScriptConfig();
  assertConfig(config);
  descargarCfdiMasivo(config.defaultOptions);
}

descargarCfdiMasivo = function(userOptions) {
  const config = getScriptConfig();
  assertConfig(config);
  const options = mergeOptions(config.defaultOptions, userOptions || {});
  validarOpciones(options);

  const token = obtenerToken(config);
  const solicitud = solicitarDescarga(config, options, token);
  const requestId = solicitud.IdSolicitud;
  const status = esperarPaquetes(config, requestId, token);

  const incomeRows = [];
  const expenseRows = [];

  status.IdsPaquetes.forEach(function (paqueteId) {
    const paquete = descargarPaquete(config, paqueteId, token);
    const zipBlob = Utilities.newBlob(paquete.zipBytes, 'application/zip', paquete.id + '.zip');
    const blobs = Utilities.unzip(zipBlob);

    blobs.forEach(function(blob) {
        const xmlTexto = blob.getDataAsString('UTF-8');
        try {
            const parsedData = parseAndFormatCfdi(xmlTexto);
            if (parsedData.tipoDeComprobante === 'I') {
                incomeRows.push(parsedData.rowData);
            } else if (parsedData.tipoDeComprobante === 'E') {
                expenseRows.push(parsedData.rowData);
            }
        } catch (e) {
            Logger.log('No se pudo interpretar ' + blob.getName() + ': ' + e.message);
        }
    });
  });

  if (incomeRows.length > 0) writeDataToSheet('XML_I', incomeRows);
  if (expenseRows.length > 0) writeDataToSheet('XML_E', expenseRows);

  return {
    requestId: requestId,
    paquetes: status.IdsPaquetes,
    hojasActualizadas: ['XML_I', 'XML_E'],
  };
}


function mergeOptions(base, override) { /* ... unchanged ... */ }
function validarOpciones(options) { /* ... unchanged ... */ }
function obtenerToken(config) { /* ... unchanged ... */ }
function solicitarDescarga(config, options, token) { /* ... unchanged ... */ }
function esperarPaquetes(config, requestId, token) { /* ... unchanged ... */ }
function verificarSolicitud(config, requestId, token) { /* ... unchanged ... */ }
function descargarPaquete(config, packageId, token) { /* ... unchanged ... */ }
function llamarSat(config, request) { /* ... unchanged ... */ }
function construirSolicitudAutenticacion(config) { /* ... unchanged ... */ }
function construirSignedInfoAutenticacion(digestValue) { /* ... unchanged ... */ }
function construirSolicitudDescarga(config, options) { /* ... unchanged ... */ }
function construirAtributosSolicitud(config, options) { /* ... unchanged ... */ }
function construirElementosHijoSolicitud(options) { /* ... unchanged ... */ }
function construirSignedInfoPeticion(digestValue) { /* ... unchanged ... */ }
function construirKeyInfo(config) { /* ... unchanged ... */ }
function construirVerificacion(config, requestId) { /* ... unchanged ... */ }
function construirDescarga(config, packageId) { /* ... unchanged ... */ }
function formatoFechaIso(fecha) { /* ... unchanged ... */ }
function sha1DigestBase64(contenido) { /* ... unchanged ... */ }
function firmarSha1(contenido, privateKeyPem) { /* ... unchanged ... */ }
function limpiarCertificado(pem) { /* ... unchanged ... */ }
function escaparXml(valor) { /* ... unchanged ... */ }
function buscarDescendiente(elemento, nombreLocal) { /* ... unchanged ... */ }
function normalizarSerial(serial) { /* ... unchanged ... */ }
// For brevity, the full unchanged functions are omitted.
// I will only include the full code for the functions I am creating or refactoring.
// The SAT communication logic remains the same.
mergeOptions = function(base, override) {
  const resultado = Object.assign({}, base);
  Object.keys(override).forEach(function (key) {
    if (override[key] !== undefined && override[key] !== null && override[key] !== '') {
      resultado[key] = override[key];
    }
  });
  if (override && Array.isArray(override.rfcReceptores)) {
    resultado.rfcReceptores = override.rfcReceptores.slice();
  }
  return resultado;
};
validarOpciones = function(options) {
  if (!options.startDate || !options.endDate) {
    throw new Error('Debes proporcionar startDate y endDate en formato YYYY-MM-DD');
  }
  if (!['emitidos', 'recibidos'].includes(options.tipoConsulta.toLowerCase())) {
    throw new Error('tipoConsulta debe ser "emitidos" o "recibidos"');
  }
};
obtenerToken = function(config) {
  const cacheKey = 'sat_token_' + config.environment;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    const data = JSON.parse(cached);
    if (Date.now() < data.expiresAt - 30000) return data.token;
  }
  const envelope = construirSolicitudAutenticacion(config);
  const response = llamarSat(config, { url: SAT_ENDPOINTS[config.environment].autenticacion, soapAction: SAT_ENDPOINTS[config.environment].soapAction.autentica, payload: envelope, authorization: null });
  const parsed = XmlService.parse(response);
  const token = buscarDescendiente(parsed.getRootElement(), 'AutenticaResult').getText();
  const expiresAt = Date.parse(buscarDescendiente(parsed.getRootElement(), 'Expires').getText());
  cache.put(cacheKey, JSON.stringify({ token: token, expiresAt: expiresAt }), 280);
  return token;
};
solicitarDescarga = function(config, options, token) {
    const payload = construirSolicitudDescarga(config, options);
    const response = llamarSat(config, { url: SAT_ENDPOINTS[config.environment].solicita, soapAction: options.tipoConsulta.toLowerCase() === 'emitidos' ? SAT_ENDPOINTS[config.environment].soapAction.solicitaEmitidos : SAT_ENDPOINTS[config.environment].soapAction.solicitaRecibidos, payload: payload, authorization: token });
    const resultado = buscarDescendiente(parsed.getRootElement(), 'SolicitaDescargaResult');
    const atributos = extraerAtributos(resultado);
    if (atributos.CodEstatus !== '5000') throw new Error('SAT rechazó la solicitud: ' + atributos.Mensaje);
    return atributos;
};
// And so on for all the other SAT-related functions...
// The key is that the parsing and writing logic is now centralized.
