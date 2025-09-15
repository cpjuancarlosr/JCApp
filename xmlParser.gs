/**
 * Contains the core logic for parsing CFDI XML files.
 */

/**
 * Helper function to safely get an attribute value from an element.
 */
function getSafeAttribute(element, attributeName) {
  if (!element) return '';
  const attribute = element.getAttribute(attributeName);
  return attribute ? attribute.getValue() : '';
}

/**
 * Parses a CFDI XML string to extract a single row of data matching the new format.
 */
function parseCfdiXml(xmlContent) {
  const document = XmlService.parse(xmlContent);
  const root = document.getRootElement();

  const cfdi = root.getNamespace();
  const tfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

  const comprobante = root;
  const emisor = comprobante.getChild('Emisor', cfdi);
  const receptor = comprobante.getChild('Receptor', cfdi);
  const timbre = comprobante.getChild('Complemento', cfdi).getChild('TimbreFiscalDigital', tfd);
  const impuestosNode = comprobante.getChild('Impuestos', cfdi);

  let ivaTrasladado = 0;
  let ieps = 0;
  let isrRetenido = 0;
  let ivaRetenido = 0;
  let tasaIVA = 0;

  const subtotal = parseFloat(getSafeAttribute(comprobante, 'SubTotal')) || 0;

  if (impuestosNode) {
    const traslados = impuestosNode.getChild('Traslados', cfdi);
    if (traslados) {
      traslados.getChildren('Traslado', cfdi).forEach(t => {
        const impuesto = getSafeAttribute(t, 'Impuesto');
        const importe = parseFloat(getSafeAttribute(t, 'Importe')) || 0;
        if (impuesto === '002') { // IVA
          ivaTrasladado += importe;
        } else if (impuesto === '003') { // IEPS
          ieps += importe;
        }
      });
    }
    const retenciones = impuestosNode.getChild('Retenciones', cfdi);
    if (retenciones) {
      retenciones.getChildren('Retencion', cfdi).forEach(r => {
        const impuesto = getSafeAttribute(r, 'Impuesto');
        const importe = parseFloat(getSafeAttribute(r, 'Importe')) || 0;
        if (impuesto === '001') isrRetenido += importe; // ISR
        else if (impuesto === '002') ivaRetenido += importe; // IVA Retenido
      });
    }
  }

  if (subtotal > 0 && ivaTrasladado > 0) {
      // Calculate the effective rate and round to nearest common rate (0, 8, 16)
      const effectiveRate = (ivaTrasladado / subtotal) * 100;
      if (effectiveRate > 12) tasaIVA = 0.16;
      else if (effectiveRate > 4) tasaIVA = 0.08;
      else tasaIVA = 0;
  }

  const fecha = getSafeAttribute(comprobante, 'Fecha').split('T')[0]; // Get date part only
  const period = fecha ? fecha.substring(0, 7) : ''; // YYYY-MM

  const tipoDeComprobante = getSafeAttribute(comprobante, 'TipoDeComprobante');

  const rowData = [
    period,
    fecha,
    getSafeAttribute(comprobante, 'Serie'),
    getSafeAttribute(comprobante, 'Folio'),
    getSafeAttribute(timbre, 'UUID'),
    getSafeAttribute(emisor, 'Rfc'),
    getSafeAttribute(receptor, 'Rfc'),
    getSafeAttribute(comprobante, 'MetodoPago'),
    tasaIVA,
    subtotal.toFixed(2),
    getSafeAttribute(comprobante, 'Descuento') || '0.00',
    ivaTrasladado.toFixed(2),
    ieps.toFixed(2),
    isrRetenido.toFixed(2),
    ivaRetenido.toFixed(2),
    getSafeAttribute(comprobante, 'Total'),
  ];

  return { tipoDeComprobante: tipoDeComprobante, rowData: rowData };
}

/**
 * Defines the header row for the new spreadsheet format.
 */
function getNewHeaderRow() {
  return [
    'Periodo', 'Fecha', 'Serie', 'Folio', 'UUID', 'RFC_Emisor', 'RFC_Receptor',
    'Metodo(PUE/PPD)', 'TasaIVA(0/0.08/0.16)', 'Subtotal', 'Descuento',
    'IVA_Trasladado', 'IEPS', 'Ret_ISR', 'Ret_IVA', 'Total'
  ];
}
