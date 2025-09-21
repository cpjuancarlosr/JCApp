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
  PRODUCTION: {
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
  },
  TEST: {
    autenticacion: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/Autenticacion/Autenticacion.svc',
    solicita: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/SolicitaDescargaService.svc',
    verifica: 'https://pruebassolicituddescargamasivasolicitud.clouda.sat.gob.mx/VerificaSolicitudDescargaService.svc',
    descarga: 'https://pruebassolicituddescargamasiva.clouda.sat.gob.mx/DescargaMasivaService.svc',
    soapAction: {
      autentica: 'http://DescargaMasivaTerceros.gob.mx/IAutenticacion/Autentica',
      solicitaEmitidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaEmitidos',
      solicitaRecibidos: 'http://DescargaMasivaTerceros.sat.gob.mx/ISolicitaDescargaService/SolicitaDescargaRecibidos',
      verifica: 'http://DescargaMasivaTerceros.sat.gob.mx/IVerificaSolicitudDescargaService/VerificaSolicitudDescarga',
      descarga: 'http://DescargaMasivaTerceros.sat.gob.mx/IDescargaMasivaTercerosService/Descargar',
    },
  },
};

const DEFAULT_OPTIONS = {
  startDate: '',
  endDate: '',
  tipoConsulta: 'emitidos',
  estadoComprobante: 'Vigente',
  tipoSolicitud: 'CFDI',
  tipoComprobante: '',
  rfcEmisor: '',
  rfcReceptor: '',
  rfcReceptores: [],
  rfcACuentaTerceros: '',
  complemento: '',
  skipDriveUpload: false,
};

function getScriptConfig() {
  const props = PropertiesService.getScriptProperties();
  const cfg = {
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
    defaultOptions: props.getProperty('SAT_DEFAULT_OPTIONS'),
  };
  if (cfg.defaultOptions) {
    try {
      const parsed = JSON.parse(cfg.defaultOptions);
      cfg.defaultOptions = Object.assign({}, DEFAULT_OPTIONS, parsed);
    } catch (err) {
      throw new Error('SAT_DEFAULT_OPTIONS tiene JSON inválido: ' + err.message);
    }
  } else {
    cfg.defaultOptions = DEFAULT_OPTIONS;
  }
  return cfg;
}

function assertConfig(config) {
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
// SAT DOWNLOAD WORKFLOW
// =================================================================

function runDescargaMasivaCfdi() {
  const config = getScriptConfig();
  assertConfig(config);
  descargarCfdiMasivo(config.defaultOptions);
}

function descargarCfdiMasivo(userOptions) {
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

  SpreadsheetApp.getUi().alert('Descarga completada', `Se procesaron ${status.IdsPaquetes.length} paquetes.`, SpreadsheetApp.getUi().ButtonSet.OK);

  return {
    requestId: requestId,
    paquetes: status.IdsPaquetes,
  };
}


function mergeOptions(base, override) {
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
}

function validarOpciones(options) {
  if (!options.startDate || !options.endDate) {
    throw new Error('Debes proporcionar startDate y endDate en formato YYYY-MM-DD');
  }
  if (!['emitidos', 'recibidos'].includes(options.tipoConsulta.toLowerCase())) {
    throw new Error('tipoConsulta debe ser "emitidos" o "recibidos"');
  }
}

function obtenerToken(config) {
  const cacheKey = 'sat_token_' + config.environment;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) {
    const data = JSON.parse(cached);
    if (Date.now() < data.expiresAt - 30000) {
      return data.token;
    }
  }
  const envelope = construirSolicitudAutenticacion(config);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].autenticacion,
    soapAction: SAT_ENDPOINTS[config.environment].soapAction.autentica,
    payload: envelope,
    authorization: null,
  });
  const parsed = XmlService.parse(response);
  const namespace = XmlService.getNamespace('s', SAT_NS.SOAP);
  const body = parsed.getRootElement().getChild('Body', namespace);
  const header = parsed.getRootElement().getChild('Header', namespace);
  if (!body) {
    throw new Error('Respuesta de Autenticación inválida');
  }
  const authResult = buscarDescendiente(body, 'AutenticaResult');
  if (!authResult) {
    throw new Error('No se encontró AutenticaResult en la respuesta de autenticación');
  }
  const token = authResult.getText();
  const timestampNode = header ? buscarDescendiente(header, 'Timestamp') : null;
  var expiresAt = Date.now() + 5 * 60 * 1000;
  if (timestampNode) {
    const expiresNode = buscarDescendiente(timestampNode, 'Expires');
    if (expiresNode) {
      expiresAt = Date.parse(expiresNode.getText()) || expiresAt;
    }
  }
  cache.put(cacheKey, JSON.stringify({ token: token, expiresAt: expiresAt }), 280);
  return token;
}

function solicitarDescarga(config, options, token) {
  const payload = construirSolicitudDescarga(config, options);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].solicita,
    soapAction: options.tipoConsulta.toLowerCase() === 'emitidos'
      ? SAT_ENDPOINTS[config.environment].soapAction.solicitaEmitidos
      : SAT_ENDPOINTS[config.environment].soapAction.solicitaRecibidos,
    payload: payload,
    authorization: token,
  });
  const parsed = XmlService.parse(response);
  const resultado = buscarDescendiente(parsed.getRootElement(), 'SolicitaDescargaResult');
  if (!resultado) {
      throw new Error('No se pudo interpretar la respuesta de solicitud de descarga. Respuesta: ' + response);
  }
  const atributos = extraerAtributos(resultado);
  if (atributos.CodEstatus && atributos.CodEstatus !== '5000') {
    throw new Error('SAT rechazó la solicitud: ' + atributos.CodEstatus + ' ' + (atributos.Mensaje || ''));
  }
  if (!atributos.IdSolicitud) {
    throw new Error('La respuesta de solicitud no incluye IdSolicitud');
  }
  return atributos;
}

function esperarPaquetes(config, requestId, token) {
  const deadline = Date.now() + config.maxWaitMinutes * 60 * 1000;
  while (true) {
    const status = verificarSolicitud(config, requestId, token);
    const estadoSolicitud = parseInt(status.EstadoSolicitud, 10);
    if (estadoSolicitud === 3) { // Terminado
      if (!status.IdsPaquetes || status.IdsPaquetes.length === 0) {
        throw new Error('La solicitud se marcó terminada pero sin paquetes disponibles');
      }
      return status;
    }
    if (estadoSolicitud === 2) { // En Proceso
       // Continue
    } else if ([4, 5, 6].indexOf(estadoSolicitud) !== -1) { // Error, Rechazado, etc.
      throw new Error('La solicitud ' + requestId + ' falló con estado ' + estadoSolicitud + ': ' + (status.Mensaje || ''));
    }
    if (Date.now() > deadline) {
      throw new Error('La solicitud ' + requestId + ' no concluyó en el tiempo máximo configurado');
    }
    Utilities.sleep(Math.max(1, config.pollIntervalSeconds) * 1000);
  }
}

function verificarSolicitud(config, requestId, token) {
  const payload = construirVerificacion(config, requestId);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].verifica,
    soapAction: SAT_ENDPOINTS[config.environment].soapAction.verifica,
    payload: payload,
    authorization: token,
  });
  const parsed = XmlService.parse(response);
  const resultado = buscarDescendiente(parsed.getRootElement(), 'VerificaSolicitudDescargaResult');
  if (!resultado) {
    throw new Error('Respuesta de verificación inválida');
  }
  const atributos = extraerAtributos(resultado);
  const ids = resultado.getChildren('IdsPaquetes', resultado.getNamespace()).map(function (node) {
    return node.getText();
  });
  atributos.IdsPaquetes = ids;
  return atributos;
}

function descargarPaquete(config, packageId, token) {
  const payload = construirDescarga(config, packageId);
  const response = llamarSat(config, {
    url: SAT_ENDPOINTS[config.environment].descarga,
    soapAction: SAT_ENDPOINTS[config.environment].soapAction.descarga,
    payload: payload,
    authorization: token,
  });
  const parsed = XmlService.parse(response);
  const paqueteNode = buscarDescendiente(parsed.getRootElement(), 'Paquete');
  if (!paqueteNode) {
    throw new Error('Respuesta de descarga sin elemento Paquete');
  }
  const zipBytes = Utilities.base64Decode(paqueteNode.getText());
  return { zipBytes: zipBytes };
}

function llamarSat(config, request) {
  const headers = {
    'Content-Type': 'text/xml; charset=utf-8',
    'Accept': 'text/xml',
    'Cache-Control': 'no-cache',
    'SOAPAction': request.soapAction,
  };
  if (request.authorization) {
    headers.Authorization = 'WRAP access_token="' + request.authorization + '"';
  }
  const respuesta = UrlFetchApp.fetch(request.url, {
    method: 'post',
    headers: headers,
    muteHttpExceptions: true,
    payload: request.payload,
  });
  const codigo = respuesta.getResponseCode();
  if (codigo < 200 || codigo >= 300) {
    throw new Error('Error HTTP ' + codigo + ' al invocar ' + request.url + ': ' + respuesta.getContentText());
  }
  return respuesta.getContentText();
}

function construirSolicitudAutenticacion(config) {
  const created = formatoFechaIso(new Date());
  const expires = formatoFechaIso(new Date(Date.now() + 5 * 60 * 1000));
  const certBase64 = limpiarCertificado(config.certificatePem);
  const timestampC14n = [
    '<u:Timestamp xmlns:u="', SAT_NS.WSU, '" u:Id="_0">',
    '<u:Created>', created, '</u:Created>',
    '<u:Expires>', expires, '</u:Expires>',
    '</u:Timestamp>'
  ].join('');
  const digestValue = sha1DigestBase64(timestampC14n);
  const signedInfo = construirSignedInfoAutenticacion(digestValue);
  const signatureValue = firmarSha1(signedInfo, config.privateKeyPem);
  const signatureXml = [
    '<Signature xmlns="', SAT_NS.DS, '">',
    signedInfo,
    '<SignatureValue>', signatureValue, '</SignatureValue>',
    '<KeyInfo><o:SecurityTokenReference xmlns:o="', SAT_NS.WSSE, '">',
    '<o:Reference ValueType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3" URI="#BinarySecurityToken"/>',
    '</o:SecurityTokenReference></KeyInfo>',
    '</Signature>'
  ].join('');
  return [
    '<s:Envelope xmlns:s="', SAT_NS.SOAP, '" xmlns:o="', SAT_NS.WSSE, '" xmlns:u="', SAT_NS.WSU, '">',
    '<s:Header>',
    '<o:Security s:mustUnderstand="1">',
    timestampC14n,
    '<o:BinarySecurityToken u:Id="BinarySecurityToken" ValueType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-x509-token-profile-1.0#X509v3" EncodingType="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-soap-message-security-1.0#Base64Binary">',
    certBase64,
    '</o:BinarySecurityToken>',
    signatureXml,
    '</o:Security>',
    '</s:Header>',
    '<s:Body><Autentica xmlns="http://DescargaMasivaTerceros.gob.mx"/></s:Body>',
    '</s:Envelope>'
  ].join('');
}

function construirSignedInfoAutenticacion(digestValue) {
  return [
    '<SignedInfo xmlns="', SAT_NS.DS, '">',
    '<CanonicalizationMethod Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/>',
    '<SignatureMethod Algorithm="http://www.w3.org/2000/09/xmldsig#rsa-sha1"/>',
    '<Reference URI="#_0">',
    '<Transforms><Transform Algorithm="http://www.w3.org/2001/10/xml-exc-c14n#"/></Transforms>',
    '<DigestMethod Algorithm="http://www.w3.org/2000/09/xmldsig#sha1"/>',
    '<DigestValue>', digestValue, '</DigestValue>',
    '</Reference>',
    '</SignedInfo>'
  ].join('');
}

function construirSolicitudDescarga(config, options) {
  const nsDes = ' xmlns:des="' + SAT_NS.DES + '"';
  const atributos = construirAtributosSolicitud(config, options);
  const hijos = construirElementosHijoSolicitud(options);
  const solicitudXml = hijos
    ? ['<des:solicitud', atributos, '>', hijos, '</des:solicitud>'].join('')
    : ['<des:solicitud', atributos, '/>'].join('');
  const operacion = options.tipoConsulta.toLowerCase() === 'emitidos'
    ? 'SolicitaDescargaEmitidos'
    : 'SolicitaDescargaRecibidos';
  const contenidoSinFirma = ['<des:', operacion, nsDes, '>', solicitudXml, '</des:', operacion, '>'].join('');
  const digestValue = sha1DigestBase64(contenidoSinFirma);
  const signedInfo = construirSignedInfoPeticion(digestValue);
  const signatureValue = firmarSha1(signedInfo, config.privateKeyPem);
  const keyInfo = construirKeyInfo(config);
  const signatureXml = ['<Signature xmlns="', SAT_NS.DS, '">', signedInfo, '<SignatureValue>', signatureValue, '</SignatureValue>', keyInfo, '</Signature>'].join('');
  const contenidoConFirma = ['<des:', operacion, nsDes, '>', solicitudXml, signatureXml, '</des:', operacion, '>'].join('');
  return ['<s:Envelope xmlns:s="', SAT_NS.SOAP, '"><s:Header/>', '<s:Body>', contenidoConFirma, '</s:Body></s:Envelope>'].join('');
}

function construirAtributosSolicitud(config, options) {
  const attrs = {
    FechaInicial: options.startDate,
    FechaFinal: options.endDate,
    RfcSolicitante: config.rfc,
    TipoSolicitud: options.tipoSolicitud,
    EstadoComprobante: options.estadoComprobante || null,
    TipoComprobante: options.tipoComprobante || null,
    RfcACuentaTerceros: options.rfcACuentaTerceros || null,
    Complemento: options.complemento || null,
  };
  const tipoConsulta = (options.tipoConsulta || '').toLowerCase();
  if (tipoConsulta === 'emitidos') {
    attrs.RfcEmisor = options.rfcEmisor || config.rfc;
    if (options.rfcReceptor) {
      attrs.RfcReceptor = options.rfcReceptor;
    }
  } else {
    attrs.RfcReceptor = options.rfcReceptor || config.rfc;
    if (options.rfcEmisor) {
      attrs.RfcEmisor = options.rfcEmisor;
    }
  }
  const entries = Object.keys(attrs).filter(function (key) {
    return attrs[key] !== undefined && attrs[key] !== null && attrs[key] !== '';
  }).map(function (key) {
    return [key, attrs[key]];
  });
  entries.sort(function (a, b) {
    return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0;
  });
  return entries.map(function (entry) {
    return ' ' + entry[0] + '="' + escaparXml(entry[1]) + '"';
  }).join('');
}

function construirElementosHijoSolicitud(options) {
  const partes = [];
  const receptores = Array.isArray(options.rfcReceptores) ? options.rfcReceptores.filter(function (r) { return r; }) : [];
  if (receptores.length) {
    const nodos = receptores.map(function (rfc) {
      return '<des:RfcReceptor>' + escaparXml(rfc) + '</des:RfcReceptor>';
    }).join('');
    partes.push('<des:RfcReceptores>' + nodos + '</des:RfcReceptores>');
  }
  return partes.join('');
}

function construirSignedInfoPeticion(digestValue) {
  return [
    '<SignedInfo xmlns="', SAT_NS.DS, '">',
    '<CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>',
    '<SignatureMethod Algorithm="http://www.w3.org/2000/09/xmldsig#rsa-sha1"/>',
    '<Reference URI="">',
    '<Transforms><Transform Algorithm="http://www.w3.org/2000/09/xmldsig#enveloped-signature"/></Transforms>',
    '<DigestMethod Algorithm="http://www.w3.org/2000/09/xmldsig#sha1"/>',
    '<DigestValue>', digestValue, '</DigestValue>',
    '</Reference>',
    '</SignedInfo>'
  ].join('');
}

function construirKeyInfo(config) {
  const certificado = limpiarCertificado(config.certificatePem);
  const issuer = escaparXml(config.certificateIssuer);
  const serial = normalizarSerial(config.certificateSerial);
  return [
    '<KeyInfo><X509Data>',
    '<X509IssuerSerial>',
    '<X509IssuerName>', issuer, '</X509IssuerName>',
    '<X509SerialNumber>', serial, '</X509SerialNumber>',
    '</X509IssuerSerial>',
    '<X509Certificate>', certificado, '</X509Certificate>',
    '</X509Data></KeyInfo>'
  ].join('');
}

function construirVerificacion(config, requestId) {
  const solicitud = ['<des:VerificaSolicitudDescarga xmlns:des="', SAT_NS.DES, '">',
    '<des:solicitud RfcSolicitante="', escaparXml(config.rfc), '" IdSolicitud="', escaparXml(requestId), '"/>',
    '</des:VerificaSolicitudDescarga>'].join('');
  return ['<s:Envelope xmlns:s="', SAT_NS.SOAP, '"><s:Header/>', '<s:Body>', solicitud, '</s:Body></s:Envelope>'].join('');
}

function construirDescarga(config, packageId) {
  const solicitud = ['<des:PeticionDescargaMasivaTercerosEntrada xmlns:des="', SAT_NS.DES, '">',
    '<des:peticionDescarga RfcSolicitante="', escaparXml(config.rfc), '" IdPaquete="', escaparXml(packageId), '"/>',
    '</des:PeticionDescargaMasivaTercerosEntrada>'].join('');
  return ['<s:Envelope xmlns:s="', SAT_NS.SOAP, '"><s:Header/>', '<s:Body>', solicitud, '</s:Body></s:Envelope>'].join('');
}

function formatoFechaIso(fecha) {
  return Utilities.formatDate(fecha, 'GMT', "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
}

function sha1DigestBase64(contenido) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, contenido, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(digest);
}

function firmarSha1(contenido, privateKeyPem) {
  const firma = Utilities.computeRsaSha1Signature(contenido, privateKeyPem);
  return Utilities.base64Encode(firma);
}

function limpiarCertificado(pem) {
  return pem.replace(/-----BEGIN CERTIFICATE-----/, '')
    .replace(/-----END CERTIFICATE-----/, '')
    .replace(/\s+/g, '');
}

function escaparXml(valor) {
  return String(valor)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buscarDescendiente(elemento, nombreLocal) {
  if (elemento.getName && elemento.getName() === nombreLocal) {
    return elemento;
  }
  const hijos = elemento.getChildren();
  for (var i = 0; i < hijos.length; i++) {
    const encontrado = buscarDescendiente(hijos[i], nombreLocal);
    if (encontrado) {
      return encontrado;
    }
  }
  return null;
}

function normalizarSerial(serial) {
  if (!serial) {
    throw new Error('FIEL_CERT_SERIAL no puede estar vacío');
  }
  const limpio = serial.replace(/[^0-9A-Fa-f]/g, '');
  if (!limpio) {
    throw new Error('FIEL_CERT_SERIAL inválido');
  }
  if (/^[0-9]+$/.test(limpio)) {
    return limpio;
  }
  return BigInt('0x' + limpio).toString(10);
}

function extraerAtributos(elemento) {
    const attrs = {};
    elemento.getAttributes().forEach(function(attr) {
        attrs[attr.getName()] = attr.getValue();
    });
    return attrs;
}
