/**
 * @file polizas.gs
 * @description Módulo para la generación de pólizas contables automáticas.
 * Contiene las reglas de negocio para traducir los CFDI y movimientos bancarios
 * en asientos de diario, ingreso y egreso.
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Orquesta la generación de todas las pólizas automáticas.
 * Limpia la hoja de pólizas y la regenera a partir de los datos de la hoja XML.
 */
function generatePolicies() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmación', 'Este proceso borrará las pólizas existentes y las generará de nuevo a partir de la hoja "XML". ¿Desea continuar?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const xmlSheet = ss.getSheetByName(SHEETS.XML);
  const polizasSheet = ss.getSheetByName(SHEETS.POLIZAS);

  // Limpiar pólizas anteriores (excepto el encabezado)
  const lastRow = polizasSheet.getLastRow();
  if (lastRow > 1) {
    polizasSheet.getRange(2, 1, lastRow - 1, polizasSheet.getMaxColumns()).clearContent();
  }

  const xmlData = xmlSheet.getDataRange().getValues();
  let policies = [];

  // Generar pólizas desde CFDI
  const cfdiPolicies = generatePoliciesFromCfdi(xmlData);
  policies = policies.concat(cfdiPolicies);

  // Escribir las nuevas pólizas en la hoja "Polizas"
  if (policies.length > 0) {
    updatePoliciesSheet(policies);
    ui.alert('Éxito', `Se generaron ${policies.length} nuevos asientos contables.`, ui.ButtonSet.OK);
  } else {
    ui.alert('Información', 'No se encontraron datos en la hoja XML para generar pólizas.', ui.ButtonSet.OK);
  }
}

/**
 * Genera pólizas a partir de los datos de CFDI.
 * @param {Array<Array<any>>} xmlData Datos de la hoja XML.
 * @returns {Array<Array<any>>} Un array de filas de pólizas.
 */
function generatePoliciesFromCfdi(xmlData) {
  const policies = [];
  const headers = xmlData.shift(); // Quitar encabezados

  xmlData.forEach(row => {
    const tipo = row[3]; // 'I', 'E', 'P'
    const subtotal = parseFloat(row[7]);
    const iva = parseFloat(row[8]);
    const total = parseFloat(row[9]);
    const uuid = row[0];
    const rfcEmisor = row[1];
    const rfcReceptor = row[2];

    if (tipo === 'I') { // Póliza de Ingreso (Venta)
      // Cargo a Clientes (105.xx)
      policies.push(createPolicyLine('Ingreso', uuid, '105.01', 'CLIENTES', total, 0));
      // Abono a Ventas (401.01)
      policies.push(createPolicyLine('Ingreso', uuid, '401.01', 'VENTAS GRAVADAS', 0, subtotal));
      // Abono a IVA por Pagar (209.01)
      policies.push(createPolicyLine('Ingreso', uuid, '209.01', 'IVA TRASLADADO', 0, iva));
    } else if (tipo === 'E') { // Póliza de Egreso (Gasto/Compra)
      // Cargo a Gastos/Compras (5xx.xx / 115.xx)
      policies.push(createPolicyLine('Egreso', uuid, '501.01', 'GASTOS GENERALES', subtotal, 0));
      // Cargo a IVA Acreditable (118.01)
      policies.push(createPolicyLine('Egreso', uuid, '118.01', 'IVA ACREDITABLE', iva, 0));
      // Abono a Proveedores (201.xx)
      policies.push(createPolicyLine('Egreso', uuid, '201.01', 'PROVEEDORES', 0, total));
    }
  });

  return policies;
}

/**
 * Genera pólizas a partir de los datos bancarios.
 * @param {Array<Array<any>>} bankData Datos de la hoja Bancos.
 * @returns {Array<Array<any>>} Un array de filas de pólizas.
 */
function generatePoliciesFromBank(bankData) {
  // Lógica para pólizas de cobro y pago, cruzando contra XML si es posible.
  return [];
}

/**
 * Crea una línea (asiento) de póliza estandarizada.
 * @param {string} type Tipo de póliza (Diario, Ingreso, Egreso).
 * @param {string} reference Referencia (UUID, folio, etc.).
 * @param {string} account Cuenta contable.
 * @param {string} concept Concepto del movimiento.
 * @param {number} debe Monto en el debe.
 * @param {number} haber Monto en el haber.
 * @returns {Array<any>} Una fila lista para insertar en la hoja "Polizas".
 */
function createPolicyLine(type, reference, account, concept, debe, haber) {
  return [
    new Date(), // Fecha de la póliza
    type,
    '', // Folio (se puede generar después)
    account,
    '', // Subcuenta (lógica de clientes/proveedores dinámicos)
    concept,
    reference,
    debe,
    haber,
    reference, // UUID
    'Auto' // Estatus
  ];
}

/**
 * Escribe las pólizas generadas en la hoja correspondiente.
 * @param {Array<Array<any>>} policies Las filas de pólizas a insertar.
 */
function updatePoliciesSheet(policies) {
  if (!policies || policies.length === 0) {
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.POLIZAS);
  const startRow = sheet.getLastRow() + 1;

  sheet.getRange(startRow, 1, policies.length, policies[0].length).setValues(policies);
  SpreadsheetApp.getUi().alert(`Se generaron y agregaron ${policies.length} asientos a la hoja "Polizas".`);
}
