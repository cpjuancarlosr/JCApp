/**
 * Contains the core logic for parsing CFDI XML files.
 */

/**
 * Helper function to safely get an attribute value from an element.
 * @param {XmlService.Element} element The element.
 * @param {string} attributeName The name of the attribute.
 * @returns {string} The attribute value or an empty string if not found.
 */
function getSafeAttribute(element, attributeName) {
  if (!element) return '';
  const attribute = element.getAttribute(attributeName);
  return attribute ? attribute.getValue() : '';
}

/**
 * Parses a CFDI XML string (versions 3.3 and 4.0) and extracts key information.
 * This function dynamically handles the CFDI namespace.
 *
 * @param {string} xmlContent The string content of the XML file.
 * @returns {Array<Array<string>>} A 2D array representing the rows to be inserted into the sheet.
 */
function parseCfdiXml(xmlContent) {
  const document = XmlService.parse(xmlContent);
  const root = document.getRootElement();

  const cfdi = root.getNamespace();
  const tfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

  const comprobante = root;
  const emisor = comprobante.getChild('Emisor', cfdi);
  const receptor = comprobante.getChild('Receptor', cfdi);
  const conceptos = comprobante.getChild('Conceptos', cfdi).getChildren('Concepto', cfdi);
  const timbre = comprobante.getChild('Complemento', cfdi).getChild('TimbreFiscalDigital', tfd);
  const impuestosNode = comprobante.getChild('Impuestos', cfdi);

  let ivaTraslado = 0;
  let ivaRetenido = 0;
  let isrRetenido = 0;

  if (impuestosNode) {
    const traslados = impuestosNode.getChild('Traslados', cfdi);
    if (traslados) {
      traslados.getChildren('Traslado', cfdi).forEach(t => {
        if (getSafeAttribute(t, 'Impuesto') === '002') {
          ivaTraslado += parseFloat(getSafeAttribute(t, 'Importe')) || 0;
        }
      });
    }
    const retenciones = impuestosNode.getChild('Retenciones', cfdi);
    if (retenciones) {
      retenciones.getChildren('Retencion', cfdi).forEach(r => {
        const impuesto = getSafeAttribute(r, 'Impuesto');
        const importe = parseFloat(getSafeAttribute(r, 'Importe')) || 0;
        if (impuesto === '002') ivaRetenido += importe;
        else if (impuesto === '001') isrRetenido += importe;
      });
    }
  }

  const fecha = getSafeAttribute(comprobante, 'Fecha');
  const date = new Date(fecha);
  const anio = fecha ? date.getFullYear().toString() : '';
  const mes = fecha ? ('0' + (date.getMonth() + 1)).slice(-2) : '';
  const period = anio && mes ? anio + '-' + mes : '';

  const generalData = {
    PERIOD: period,
    ANIO: anio,
    MES: mes,
    SERIE: getSafeAttribute(comprobante, 'Serie'),
    FOLIO: getSafeAttribute(comprobante, 'Folio'),
    TIPODECOMPROBANTE: getSafeAttribute(comprobante, 'TipoDeComprobante'),
    FECHA: fecha,
    EMISORREGIMENFISCAL: getSafeAttribute(emisor, 'RegimenFiscal'),
    EMISORRFC: getSafeAttribute(emisor, 'Rfc'),
    EMISORNOMBRE: getSafeAttribute(emisor, 'Nombre'),
    RECEPTORRFC: getSafeAttribute(receptor, 'Rfc'),
    RECEPTORNOMBRE: getSafeAttribute(receptor, 'Nombre'),
    RECEPTORUSOCFDI: getSafeAttribute(receptor, 'UsoCFDI'),
    ESTATUS: 'Vigente',
    MONEDA: getSafeAttribute(comprobante, 'Moneda'),
    METODOPAGO: getSafeAttribute(comprobante, 'MetodoPago'),
    FORMAPAGO: getSafeAttribute(comprobante, 'FormaPago'),
    SUBTOTAL: getSafeAttribute(comprobante, 'SubTotal'),
    DESCUENTO: getSafeAttribute(comprobante, 'Descuento') || '0',
    IVATRASLADO: ivaTraslado.toFixed(2),
    IVARETENIDO: ivaRetenido.toFixed(2),
    ISRRETENIDO: isrRetenido.toFixed(2),
    TOTAL: getSafeAttribute(comprobante, 'Total'),
    URLXML: '',
    URLPDF: '',
    UUID: getSafeAttribute(timbre, 'UUID'),
    FECHACANCELACION: '',
  };

  const rows = [];
  conceptos.forEach((concepto, index) => {
    const isFirstLine = index === 0;
    const row = [
      // Descriptive data (repeated for context)
      generalData.PERIOD,
      generalData.ANIO,
      generalData.MES,
      generalData.SERIE,
      generalData.FOLIO,
      generalData.TIPODECOMPROBANTE,
      generalData.FECHA,
      generalData.EMISORREGIMENFISCAL,
      generalData.EMISORRFC,
      generalData.EMISORNOMBRE,
      generalData.RECEPTORRFC,
      generalData.RECEPTORNOMBRE,
      // Line-item specific data
      getSafeAttribute(concepto, 'ClaveProdServ'),
      // More descriptive data
      generalData.RECEPTORUSOCFDI,
      generalData.ESTATUS,
      generalData.MONEDA,
      generalData.METODOPAGO,
      generalData.FORMAPAGO,
      // Financial data (only on first line)
      isFirstLine ? generalData.SUBTOTAL : '',
      isFirstLine ? generalData.DESCUENTO : '',
      isFirstLine ? generalData.IVATRASLADO : '',
      isFirstLine ? generalData.IVARETENIDO : '',
      isFirstLine ? generalData.ISRRETENIDO : '',
      isFirstLine ? generalData.TOTAL : '',
      // Descriptive data (repeated for context)
      generalData.URLXML,
      generalData.URLPDF,
      generalData.UUID,
      generalData.FECHACANCELACION,
    ];
    rows.push(row);
  });

  return rows;
}

/**
 * Defines the header row for the spreadsheet.
 * @returns {Array<string>} The header row.
 */
function getHeaderRow() {
  return [
    'PERIOD', 'ANIO', 'MES', 'SERIE', 'FOLIO', 'TIPODECOMPROBANTE', 'FECHA',
    'EMISORREGIMENFISCAL', 'EMISORRFC', 'EMISORNOMBRE', 'RECEPTORRFC', 'RECEPTORNOMBRE',
    'CLAVEPRODSERV', 'RECEPTORUSOCFDI', 'ESTATUS', 'MONEDA', 'METODOPAGO', 'FORMAPAGO',
    'SUBTOTAL', 'DESCUENTO', 'IVATRASLADO', 'IVARETENIDO', 'ISRRETENIDO', 'TOTAL',
    'URLXML', 'URLPDF', 'UUID', 'FECHACANCELACION'
  ];
}
