/**
 * @file xml_parser.gs
 * @description Módulo para parsear (analizar) el contenido de los archivos XML (CFDI)
 * y extraer la información relevante para ser volcada en la hoja "XML".
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Parsea el contenido de un string XML de un CFDI y extrae los datos clave.
 * @param {string} xmlContent El contenido del archivo XML.
 * @returns {Array|null} Un array con los datos estructurados para una fila de la hoja "XML", o null si hay un error.
 */
function parseXml(xmlContent) {
  try {
    const document = XmlService.parse(xmlContent);
    const root = document.getRootElement();
    const nsCfdi = root.getNamespace();
    const nsTfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');

    const emisor = root.getChild('Emisor', nsCfdi);
    const receptor = root.getChild('Receptor', nsCfdi);
    const complemento = root.getChild('Complemento', nsCfdi);
    const timbre = complemento.getChild('TimbreFiscalDigital', nsTfd);

    const uuid = timbre.getAttribute('UUID').getValue();
    const emisorRfc = emisor.getAttribute('Rfc').getValue();
    const receptorRfc = receptor.getAttribute('Rfc').getValue();
    const tipo = root.getAttribute('TipoDeComprobante').getValue();
    const fecha = root.getAttribute('Fecha').getValue();
    const subtotal = root.getAttribute('SubTotal').getValue();
    const total = root.getAttribute('Total').getValue();
    const moneda = root.getAttribute('Moneda').getValue();

    // Estos campos pueden no existir y necesitan manejo de errores
    const metodoPago = root.getAttribute('MetodoPago') ? root.getAttribute('MetodoPago').getValue() : '';
    const formaPago = root.getAttribute('FormaPago') ? root.getAttribute('FormaPago').getValue() : '';
    const usoCfdi = receptor.getAttribute('UsoCFDI') ? receptor.getAttribute('UsoCFDI').getValue() : '';

    // Extraer impuestos (lógica simplificada)
    // Una implementación completa requeriría iterar sobre el nodo de Impuestos
    const impuestosNode = root.getChild('Impuestos', nsCfdi);
    let totalImpuestosTrasladados = 0;
    if (impuestosNode && impuestosNode.getAttribute('TotalImpuestosTrasladados')) {
        totalImpuestosTrasladados = impuestosNode.getAttribute('TotalImpuestosTrasladados').getValue();
    }

    // Devolver un array en el orden de las columnas de la hoja "XML"
    return [
      uuid,
      emisorRfc,
      receptorRfc,
      tipo, // I (Ingreso), E (Egreso), P (Pago)
      fecha,
      metodoPago, // PUE, PPD
      formaPago,
      subtotal,
      totalImpuestosTrasladados, // Simplificado
      total,
      usoCfdi,
      moneda,
      '' // CFDI Relacionado (placeholder)
    ];
  } catch (e) {
    Logger.log(`Error parseando XML: ${e.message}`);
    return null;
  }
}

/**
 * Escribe los datos parseados de múltiples XMLs en la hoja "XML".
 * @param {Array<Array<any>>} dataRows Un array de filas, donde cada fila es un array de datos de un CFDI.
 */
function updateXmlSheet(dataRows) {
  if (!dataRows || dataRows.length === 0) {
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.XML);
  const startRow = sheet.getLastRow() + 1;
  const numColumns = dataRows[0].length;

  sheet.getRange(startRow, 1, dataRows.length, numColumns).setValues(dataRows);
  SpreadsheetApp.getUi().alert(`Se procesaron y agregaron ${dataRows.length} CFDI a la hoja "XML".`);
}
