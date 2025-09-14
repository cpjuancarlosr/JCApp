/**
 * @file reports.gs
 * @description Módulo para la generación de reportes financieros.
 * Contiene funciones para refrescar tablas dinámicas, aplicar fórmulas matriciales
 * y calcular los KPIs que se mostrarán en el dashboard "Inicio".
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Coloca las fórmulas para los reportes financieros dinámicos en la hoja "Inicio".
 */
function setupFinancialReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const homeSheet = ss.getSheetByName(SHEETS.HOME);
  if (!homeSheet) {
    SpreadsheetApp.getUi().alert('Error', `No se encuentra la hoja "${SHEETS.HOME}". Por favor, ejecute la configuración de hojas primero.`);
    return;
  }

  // Limpiar área de reportes (ej. columnas K en adelante)
  homeSheet.getRange('K:Z').clear();

  // --- Balanza de Comprobación ---
  const balanceTitleCell = homeSheet.getRange('K1');
  balanceTitleCell.setValue('Balanza de Comprobación (Saldos)');
  formatAsTitle(balanceTitleCell.mergeTo('N1'));

  const balanceHeaderCell = homeSheet.getRange('K2');
  const balanceFormulaCell = homeSheet.getRange('K3');
  const polizasRange = `'${SHEETS.POLIZAS}'!D:I`; // Rango de Cuenta, Subcuenta, ..., Debe, Haber

  // Encabezados
  balanceHeaderCell.setValues([['Cuenta', 'Debe', 'Haber', 'Saldo Final']]);
  formatAsHeader(homeSheet.getRange('K2:N2'));

  // Fórmula QUERY para la balanza
  const balanceFormula = `=QUERY(${polizasRange}, "SELECT D, SUM(H), SUM(I), SUM(H)-SUM(I) WHERE D IS NOT NULL GROUP BY D ORDER BY D LABEL D 'Cuenta', SUM(H) 'Debe', SUM(I) 'Haber', SUM(H)-SUM(I) 'Saldo Final'")`;
  balanceFormulaCell.setFormula(balanceFormula);

  // --- Estado de Resultados ---
  const incomeTitleCell = homeSheet.getRange('P1');
  incomeTitleCell.setValue('Estado de Resultados');
  formatAsTitle(incomeTitleCell.mergeTo('R1'));

  const incomeHeaderCell = homeSheet.getRange('P2');
  const incomeFormulaCell = homeSheet.getRange('P3');

  // Encabezados
  incomeHeaderCell.setValues([['Cuenta', 'Concepto', 'Saldo']]);
  formatAsHeader(homeSheet.getRange('P2:R2'));

  // Fórmula QUERY para el estado de resultados
  // Se une con la hoja de Cuentas para traer el nombre de la cuenta.
  const incomeFormula = `=QUERY({QUERY(${polizasRange}, "SELECT D, SUM(I)-SUM(H) WHERE D >= '4' AND D < '6' GROUP BY D"), VLOOKUP(QUERY(${polizasRange}, "SELECT D WHERE D >= '4' AND D < '6' GROUP BY D LABEL D ''"), '${SHEETS.CUENTAS}'!A:B, 2, FALSE)}, "SELECT Col2, Col1, Col3 LABEL Col2 'Cuenta', Col1 'Concepto', Col3 'Saldo' FORMAT Col3 '$#,##0.00'")`;

  // Fórmula más simple sin VLOOKUP por si Cuentas no está poblada
  const simpleIncomeFormula = `=QUERY(${polizasRange}, "SELECT D, SUM(I)-SUM(H) WHERE D IS NOT NULL AND (D LIKE '4%' OR D LIKE '5%') GROUP BY D LABEL D 'Cuenta', SUM(I)-SUM(H) 'Saldo'")`;
  incomeFormulaCell.setFormula(simpleIncomeFormula);

  SpreadsheetApp.getUi().alert('Reportes Actualizados', 'Las fórmulas de los reportes dinámicos han sido colocadas en la hoja "Inicio".', SpreadsheetApp.getUi().ButtonSet.OK);
}
