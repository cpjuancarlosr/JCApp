/**
 * @file xml_parser.gs
 * @description Módulo para parsear (analizar) XML de CFDI y extraer información clave.
 *
 * @author Jules
 * @version 3.0.0
 */

function getTimbreFiscal(xmlContent) {
    try {
        const doc = XmlService.parse(xmlContent);
        const root = doc.getRootElement();
        const cfdi = root.getNamespace();
        const tfd = XmlService.getNamespace('tfd', 'http://www.sat.gob.mx/TimbreFiscalDigital');
        const complemento = root.getChild('Complemento', cfdi);
        if (!complemento) return {};
        const timbre = complemento.getChild('TimbreFiscalDigital', tfd);
        if (!timbre) return {};
        return {
            uuid: timbre.getAttribute('UUID')?.getValue(),
            fechaTimbrado: timbre.getAttribute('FechaTimbrado')?.getValue()
        };
    } catch (e) {
        Logger.log(`No se pudo obtener el Timbre Fiscal: ${e.message}`);
        return {};
    }
}

function parsearCFDI(xmlContent, timbreData) {
  try {
    const doc = XmlService.parse(xmlContent);
    const root = doc.getRootElement();
    const cfdi = root.getNamespace();

    const emisor = root.getChild('Emisor', cfdi);
    const receptor = root.getChild('Receptor', cfdi);

    let cfdiData = {
      uuid: timbreData.uuid,
      fecha: root.getAttribute('Fecha').getValue(),
      tipo: root.getAttribute('TipoDeComprobante').getValue(),
      formaPago: root.getAttribute('FormaPago')?.getValue() || '',
      metodoPago: root.getAttribute('MetodoPago')?.getValue() || '',
      subtotal: parseFloat(root.getAttribute('SubTotal').getValue()),
      total: parseFloat(root.getAttribute('Total').getValue()),
      emisorRfc: emisor.getAttribute('Rfc').getValue(),
      receptorRfc: receptor.getAttribute('Rfc').getValue(),
      ivaTrasladado: 0,
      ivaRetenido: 0,
      pagoIvaTrasladado: 0, // Campo específico para IVA en Complemento de Pago
      uuidRelacionado: null,
      claveProdServ: '',
      usoCFDI: receptor.getAttribute('UsoCFDI').getValue()
    };

    // --- Impuestos Generales (para Facturas I y E) ---
    const impuestosNode = root.getChild('Impuestos', cfdi);
    if (impuestosNode) {
        (impuestosNode.getChild('Traslados', cfdi)?.getChildren('Traslado', cfdi) || []).forEach(t => {
            if (t.getAttribute('Impuesto').getValue() === '002') cfdiData.ivaTrasladado += parseFloat(t.getAttribute('Importe').getValue());
        });
        (impuestosNode.getChild('Retenciones', cfdi)?.getChildren('Retencion', cfdi) || []).forEach(r => {
            if (r.getAttribute('Impuesto').getValue() === '002') cfdiData.ivaRetenido += parseFloat(r.getAttribute('Importe').getValue());
        });
    }

    // --- CFDI Relacionado ---
    const cfdiRelacionadosNode = root.getChild('CfdiRelacionados', cfdi);
    if (cfdiRelacionadosNode) {
      cfdiData.uuidRelacionado = cfdiRelacionadosNode.getChild('CfdiRelacionado', cfdi)?.getAttribute('UUID').getValue();
    }

    // --- Conceptos (para Facturas I y E) ---
    const primerConcepto = root.getChild('Conceptos', cfdi)?.getChild('Concepto', cfdi);
    if (primerConcepto) cfdiData.claveProdServ = primerConcepto.getAttribute('ClaveProdServ').getValue();

    // --- Lógica Específica para Complemento de Pago (Tipo 'P') ---
    if (cfdiData.tipo === 'P') {
        const complemento = root.getChild('Complemento', cfdi);
        const pago20 = XmlService.getNamespace('pago20', 'http://www.sat.gob.mx/Pagos20');
        const pagos = complemento.getChild('Pagos', pago20);
        if (pagos) {
            const pago = pagos.getChild('Pago', pago20);
            // Tomamos el UUID del primer documento relacionado en el pago, que es lo más común.
            cfdiData.uuidRelacionado = pago.getChild('DoctoRelacionado', pago20)?.getAttribute('IdDocumento').getValue();

            const impuestosP = pago.getChild('ImpuestosP', pago20);
            if (impuestosP) {
                (impuestosP.getChild('TrasladosP', pago20)?.getChildren('TrasladoP', pago20) || []).forEach(t => {
                    if (t.getAttribute('ImpuestoP').getValue() === '002') { // IVA
                        cfdiData.pagoIvaTrasladado += parseFloat(t.getAttribute('ImporteP').getValue());
                    }
                });
            }
        }
    }

    // --- Validación de Integridad (solo para facturas tipo I y E) ---
    if (cfdiData.tipo !== 'P') {
        const totalCalculado = cfdiData.subtotal + cfdiData.ivaTrasladado - cfdiData.ivaRetenido;
        if (Math.abs(totalCalculado - cfdiData.total) > 0.02) {
            Logger.log(`ERROR DE VALIDACIÓN: El total no cuadra en ${cfdiData.uuid}. Calculado: ${totalCalculado}, Declarado: ${cfdiData.total}`);
            return null;
        }
    }

    return cfdiData;
  } catch (e) {
    Logger.log(`Error crítico al parsear CFDI con UUID ${timbreData.uuid}: ${e.stack}`);
    return null;
  }
}
