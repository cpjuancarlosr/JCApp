/**
 * Contains the core logic for parsing CFDI XML files.
 */

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

  // Dynamically get the cfdi namespace from the root element. This makes it
  // compatible with both CFDI 3.3 and 4.0, as the main difference is the namespace URL.
  const cfdi = root.getNamespace();
  const tfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

  // --- Extract data ---
  const comprobante = root;
  const emisor = comprobante.getChild('Emisor', cfdi);
  const receptor = comprobante.getChild('Receptor', cfdi);
  const conceptos = comprobante.getChild('Conceptos', cfdi).getChildren('Concepto', cfdi);
  const timbre = comprobante.getChild('Complemento', cfdi).getChild('TimbreFiscalDigital', tfd);

  // General Data
  const generalData = {
    serie: comprobante.getAttribute('Serie') ? comprobante.getAttribute('Serie').getValue() : '',
    folio: comprobante.getAttribute('Folio') ? comprobante.getAttribute('Folio').getValue() : '',
    fecha: comprobante.getAttribute('Fecha').getValue(),
    formaPago: comprobante.getAttribute('FormaPago') ? comprobante.getAttribute('FormaPago').getValue() : '',
    metodoPago: comprobante.getAttribute('MetodoPago') ? comprobante.getAttribute('MetodoPago').getValue() : '',
    moneda: comprobante.getAttribute('Moneda').getValue(),
    subTotal: comprobante.getAttribute('SubTotal').getValue(),
    total: comprobante.getAttribute('Total').getValue(),
    tipoDeComprobante: comprobante.getAttribute('TipoDeComprobante').getValue(),
    lugarExpedicion: comprobante.getAttribute('LugarExpedicion').getValue(),
    emisorRfc: emisor.getAttribute('Rfc').getValue(),
    emisorNombre: emisor.getAttribute('Nombre').getValue(),
    receptorRfc: receptor.getAttribute('Rfc').getValue(),
    receptorNombre: receptor.getAttribute('Nombre').getValue(),
    uuid: timbre.getAttribute('UUID').getValue(),
    fechaTimbrado: timbre.getAttribute('FechaTimbrado').getValue()
  };

  const rows = [];
  // Loop through each "Concepto"
  conceptos.forEach(concepto => {
    const row = [
      generalData.serie,
      generalData.folio,
      generalData.fecha,
      generalData.formaPago,
      generalData.metodoPago,
      generalData.moneda,
      generalData.subTotal,
      generalData.total,
      generalData.tipoDeComprobante,
      generalData.lugarExpedicion,
      generalData.emisorRfc,
      generalData.emisorNombre,
      generalData.receptorRfc,
      generalData.receptorNombre,
      concepto.getAttribute('ClaveProdServ').getValue(),
      concepto.getAttribute('Cantidad').getValue(),
      concepto.getAttribute('ClaveUnidad').getValue(),
      concepto.getAttribute('Descripcion').getValue(),
      concepto.getAttribute('ValorUnitario').getValue(),
      concepto.getAttribute('Importe').getValue(),
      generalData.uuid,
      generalData.fechaTimbrado,
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
    'Serie', 'Folio', 'Fecha', 'FormaPago', 'MetodoPago', 'Moneda', 'SubTotal', 'Total',
    'TipoDeComprobante', 'LugarExpedicion', 'Emisor RFC', 'Emisor Nombre', 'Receptor RFC', 'Receptor Nombre',
    'Concepto ClaveProdServ', 'Concepto Cantidad', 'Concepto ClaveUnidad', 'Concepto Descripcion',
    'Concepto ValorUnitario', 'Concepto Importe', 'UUID', 'FechaTimbrado'
  ];
}
