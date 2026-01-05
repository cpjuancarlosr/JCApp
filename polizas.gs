/**
 * @file polizas.gs
 * @description Módulo para la generación de pólizas contables a partir de datos de CFDI.
 *
 * @author Jules
 * @version 3.3.0
 */

const CUENTAS_CLAVE = {
  IVA_ACREDITABLE_PAGADO: "118.01",
  IVA_ACREDITABLE_PENDIENTE: "118.02",
  IVA_TRASLADADO_COBRADO: "208.01",
  IVA_TRASLADADO_PENDIENTE: "208.02",
  CLIENTES: "105.01",
  PROVEEDORES: "201.01",
  BANCOS: "102.01",
  DEVOLUCION_VENTA: "402.01", // Cuenta para devoluciones sobre venta
  DEVOLUCION_COMPRA: "502.01" // Cuenta para devoluciones sobre compra
};

let cacheMapeoDeCuentas = null;

/**
 * Genera la póliza contable para un CFDI, delegando al tipo de comprobante correcto.
 * @param {object} cfdiData - Datos parseados del CFDI.
 * @param {string} nuestroRfc - RFC de la empresa que opera el sistema.
 * @return {Array<Array>} La póliza como un array de asientos.
 */
function generarPoliza(cfdiData, nuestroRfc) {
  const { tipo, receptorRfc } = cfdiData;
  const esGasto = receptorRfc.toUpperCase() === nuestroRfc.toUpperCase();

  switch (tipo) {
    case 'I':
      return esGasto ? generarPolizaGasto(cfdiData) : generarPolizaIngreso(cfdiData);
    case 'E':
      return esGasto ? generarPolizaDevolucionCompra(cfdiData) : generarPolizaDevolucionVenta(cfdiData);
    case 'P':
      return esGasto ? generarPolizaPagoAProveedor(cfdiData) : generarPolizaCobroACliente(cfdiData);
    default:
      Logger.log(`Tipo de comprobante '${tipo}' no soportado.`);
      return [];
  }
}

function generarPolizaIngreso(cfdiData) {
  const { metodoPago, total, subtotal, ivaTrasladado, uuid, claveProdServ, usoCFDI } = cfdiData;
  const poliza = [], fecha = cfdiData.fecha.substring(0, 10), polizaId = `DI-${Date.now()}`;
  const cuentaIngreso = obtenerCuentaPorMapeo(claveProdServ, usoCFDI, 'INGRESO');
  const concepto = `Venta s/factura ${uuid.substring(0, 8)}`;

  if (metodoPago === 'PUE') {
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.BANCOS, concepto, total, 0]);
    poliza.push([fecha, polizaId, uuid, cuentaIngreso, concepto, 0, subtotal]);
    if (ivaTrasladado > 0) poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_TRASLADADO_COBRADO, concepto, 0, ivaTrasladado]);
  } else {
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.CLIENTES, concepto, total, 0]);
    poliza.push([fecha, polizaId, uuid, cuentaIngreso, concepto, 0, subtotal]);
    if (ivaTrasladado > 0) poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_TRASLADADO_PENDIENTE, concepto, 0, ivaTrasladado]);
  }
  return poliza;
}

function generarPolizaGasto(cfdiData) {
  const { metodoPago, total, subtotal, ivaTrasladado, uuid, claveProdServ, usoCFDI } = cfdiData;
  const poliza = [], fecha = cfdiData.fecha.substring(0, 10), polizaId = `EG-${Date.now()}`;
  const cuentaGasto = obtenerCuentaPorMapeo(claveProdServ, usoCFDI, 'GASTO');
  const concepto = `Compra s/factura ${uuid.substring(0, 8)}`;

  if (metodoPago === 'PUE') {
    poliza.push([fecha, polizaId, uuid, cuentaGasto, concepto, subtotal, 0]);
    if (ivaTrasladado > 0) poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_ACREDITABLE_PAGADO, concepto, ivaTrasladado, 0]);
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.BANCOS, concepto, 0, total]);
  } else {
    poliza.push([fecha, polizaId, uuid, cuentaGasto, concepto, subtotal, 0]);
    if (ivaTrasladado > 0) poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_ACREDITABLE_PENDIENTE, concepto, ivaTrasladado, 0]);
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.PROVEEDORES, concepto, 0, total]);
  }
  return poliza;
}

function generarPolizaDevolucionVenta(cfdiData) {
  const { total, subtotal, ivaTrasladado, uuid } = cfdiData;
  const poliza = [], fecha = cfdiData.fecha.substring(0, 10), polizaId = `DI-${Date.now()}`;
  const concepto = `Devolución s/venta ${uuid.substring(0, 8)}`;

  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.DEVOLUCION_VENTA, concepto, subtotal, 0]);
  if (ivaTrasladado > 0) poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_TRASLADADO_PENDIENTE, concepto, ivaTrasladado, 0]);
  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.CLIENTES, concepto, 0, total]);
  return poliza;
}

function generarPolizaDevolucionCompra(cfdiData) {
  const { total, subtotal, ivaTrasladado, uuid } = cfdiData;
  const poliza = [], fecha = cfdiData.fecha.substring(0, 10), polizaId = `EG-${Date.now()}`;
  const concepto = `Devolución s/compra ${uuid.substring(0, 8)}`;

  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.PROVEEDORES, concepto, total, 0]);
  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.DEVOLUCION_COMPRA, concepto, 0, subtotal]);
  if (ivaTrasladado > 0) poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_ACREDITABLE_PENDIENTE, concepto, 0, ivaTrasladado]);
  return poliza;
}

function generarPolizaCobroACliente(cfdiData) {
  const { total, uuid, uuidRelacionado, pagoIvaTrasladado } = cfdiData;
  const poliza = [], fecha = cfdiData.fecha.substring(0, 10), polizaId = `IN-${Date.now()}`;
  const concepto = `Cobro s/factura ${uuidRelacionado ? uuidRelacionado.substring(0, 8) : ''}`;

  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.BANCOS, concepto, total, 0]);
  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.CLIENTES, concepto, 0, total]);
  if (pagoIvaTrasladado > 0) {
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_TRASLADADO_PENDIENTE, "Reclasificación IVA", pagoIvaTrasladado, 0]);
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_TRASLADADO_COBRADO, "Reclasificación IVA", 0, pagoIvaTrasladado]);
  }
  return poliza;
}

function generarPolizaPagoAProveedor(cfdiData) {
  const { total, uuid, uuidRelacionado, pagoIvaTrasladado } = cfdiData;
  const poliza = [], fecha = cfdiData.fecha.substring(0, 10), polizaId = `EG-${Date.now()}`;
  const concepto = `Pago s/factura ${uuidRelacionado ? uuidRelacionado.substring(0, 8) : ''}`;

  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.PROVEEDORES, concepto, total, 0]);
  poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.BANCOS, concepto, 0, total]);
  if (pagoIvaTrasladado > 0) {
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_ACREDITABLE_PAGADO, "Reclasificación IVA", pagoIvaTrasladado, 0]);
    poliza.push([fecha, polizaId, uuid, CUENTAS_CLAVE.IVA_ACREDITABLE_PENDIENTE, "Reclasificación IVA", 0, pagoIvaTrasladado]);
  }
  return poliza;
}

function obtenerCuentaPorMapeo(claveProdServ, usoCFDI, tipoPoliza) {
  if (cacheMapeoDeCuentas === null) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mapeo_ProdServ_Cuenta');
    const lastRow = sheet.getLastRow();
    cacheMapeoDeCuentas = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    Logger.log("Caché de mapeo de cuentas inicializada.");
  }
  for (const row of cacheMapeoDeCuentas) if (row[0] == claveProdServ && row[1] == usoCFDI) return row[2];
  for (const row of cacheMapeoDeCuentas) if (row[0] == claveProdServ) return row[2];
  return tipoPoliza === 'INGRESO' ? '401.01' : '601.01';
}

function registrarPolizaEnSheet(poliza) {
  if (poliza && poliza.length > 0) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pólizas');
    sheet.getRange(sheet.getLastRow() + 1, 1, poliza.length, poliza[0].length).setValues(poliza);
  }
}

function registrarEnLog(uuid) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log_Procesados').appendRow([uuid, new Date()]);
}
