/**
 * @file reports.gs
 * @description Módulo para la generación de reportes financieros.
 *
 * @author Jules
 * @version 3.0.0
 */

// --- Funciones Públicas (llamadas desde la UI) ---

function generarBalanzaDeComprobacion() {
  const periodo = getConfiguracion_('Periodo Contable Actual (YYYY-MM)');
  SpreadsheetApp.getActiveSpreadsheet().toast(`Generando Balanza para ${periodo}...`, 'Proceso', -1);
  const catalogo = leerHoja_('Catálogo de Cuentas'), polizas = leerHoja_('Pólizas');
  const balanzaData = procesarDatosParaBalanza_(catalogo, polizas, periodo);
  escribirReporteBalanza_(balanzaData);
  SpreadsheetApp.getActiveSpreadsheet().toast('Balanza generada con éxito.');
  return { status: 'success', message: `Balanza de Comprobación generada para ${periodo}.` };
}

function generarBalanceGeneral() {
  const periodo = getConfiguracion_('Periodo Contable Actual (YYYY-MM)');
  SpreadsheetApp.getActiveSpreadsheet().toast(`Generando Balance General para ${periodo}...`, 'Proceso', -1);
  generarBalanzaDeComprobacion();
  const catalogo = leerHoja_('Catálogo de Cuentas'), balanza = leerHoja_('Balanza de Comprobación');
  const reporteData = construirBalanceGeneral_(catalogo, balanza, 0); // Asume resultado del ejercicio 0 por ahora
  escribirReporteGeneral_('Balance General', reporteData, ['A', 'B', 'C']);
  SpreadsheetApp.getActiveSpreadsheet().toast('Balance General generado.');
  return { status: 'success', message: `Balance General generado para ${periodo}.` };
}

function generarEstadoDeResultados() {
  const periodo = getConfiguracion_('Periodo Contable Actual (YYYY-MM)');
  SpreadsheetApp.getActiveSpreadsheet().toast(`Generando Estado de Resultados para ${periodo}...`, 'Proceso', -1);
  generarBalanzaDeComprobacion();
  const catalogo = leerHoja_('Catálogo de Cuentas'), balanza = leerHoja_('Balanza de Comprobación');
  const [reporteData, resultadoNeto] = construirEstadoDeResultados_(catalogo, balanza);
  escribirReporteGeneral_('Estado de Resultados', reporteData, ['A', 'B']);

  // Re-generar Balance General con el resultado del ejercicio correcto
  const reporteBalance = construirBalanceGeneral_(catalogo, balanza, resultadoNeto);
  escribirReporteGeneral_('Balance General', reporteBalance, ['A', 'B', 'C']);

  SpreadsheetApp.getActiveSpreadsheet().toast('Reportes financieros actualizados.');
  return { status: 'success', message: `Estado de Resultados (y Balance) generado para ${periodo}.` };
}

// --- Lógica de Procesamiento ---

function procesarDatosParaBalanza_(catalogo, polizas, periodo) {
  const cuentasMap = new Map();
  catalogo.slice(1).forEach(row => {
    cuentasMap.set(row[0].toString(), { n: row[1], sI: 0, d: 0, h: 0, nat: row[4] });
  });
  polizas.slice(1).forEach(row => {
    if (row[0].toISOString().substring(0, 7) === periodo) {
      const cta = cuentasMap.get(row[3].toString());
      if (cta) { cta.d += parseFloat(row[5] || 0); cta.h += parseFloat(row[6] || 0); }
    }
  });
  return Array.from(cuentasMap)
    .filter(([_, data]) => data.d !== 0 || data.h !== 0)
    .map(([num, data]) => [num, data.n, 0, data.d, data.h, data.nat === 'Deudora' ? data.d - data.h : data.h - data.d]);
}

function construirEstadoDeResultados_(catalogo, balanza) {
    const cuentasInfo = new Map(catalogo.slice(1).map(r => [r[0].toString(), { tipo: r[2] }]));
    const saldos = new Map(balanza.slice(1).map(r => [r[0].toString(), { nombre: r[1], saldo: r[5] }]));

    let totalIngresos = 0, totalCostos = 0, totalGastos = 0;
    const reporte = [['ESTADO DE RESULTADOS', '']];

    reporte.push(['Ingresos', '']);
    saldos.forEach((data, cuenta) => {
        if (cuentasInfo.get(cuenta)?.tipo === 'INGRESO') {
            reporte.push([`  ${data.nombre}`, data.saldo]);
            totalIngresos += data.saldo;
        }
    });
    reporte.push(['TOTAL INGRESOS', totalIngresos]);
    reporte.push(['', '']);

    reporte.push(['Costos', '']);
    saldos.forEach((data, cuenta) => {
        if (cuentasInfo.get(cuenta)?.tipo === 'COSTO') {
            reporte.push([`  ${data.nombre}`, data.saldo]);
            totalCostos += data.saldo;
        }
    });
    reporte.push(['TOTAL COSTOS', totalCostos]);
    reporte.push(['UTILIDAD BRUTA', totalIngresos - totalCostos]);
    reporte.push(['', '']);

    reporte.push(['Gastos', '']);
     saldos.forEach((data, cuenta) => {
        if (cuentasInfo.get(cuenta)?.tipo === 'GASTO') {
            reporte.push([`  ${data.nombre}`, data.saldo]);
            totalGastos += data.saldo;
        }
    });
    reporte.push(['TOTAL GASTOS', totalGastos]);
    reporte.push(['', '']);

    const resultadoNeto = totalIngresos - totalCostos - totalGastos;
    reporte.push(['RESULTADO NETO DEL EJERCICIO', resultadoNeto]);

    return [reporte, resultadoNeto];
}

function construirBalanceGeneral_(catalogo, balanza, resultadoDelEjercicio) {
    const cuentasInfo = new Map(catalogo.slice(1).map(r => [r[0].toString(), { t: r[2], st: r[3] }]));
    const saldos = new Map(balanza.slice(1).map(r => [r[0].toString(), { n: r[1], s: r[5] }]));
    const estructura = { 'ACTIVO': {}, 'PASIVO': {}, 'CAPITAL': {} };
    let totales = { 'ACTIVO': 0, 'PASIVO': 0, 'CAPITAL': 0 };

    saldos.forEach((data, cta) => {
        const info = cuentasInfo.get(cta);
        if (info && estructura[info.t]) {
            if (!estructura[info.t][info.st]) estructura[info.t][info.st] = [];
            estructura[info.t][info.st].push([`   ${cta}`, data.n, data.s]);
            totales[info.t] += data.s;
        }
    });

    let reporte = [];
    ['ACTIVO', 'PASIVO', 'CAPITAL'].forEach(tipo => {
        reporte.push([tipo, '', '']);
        for (const subtipo in estructura[tipo]) {
            reporte.push([` ${subtipo}`, '', estructura[tipo][subtipo].reduce((acc, curr) => acc + curr[2], 0)]);
            reporte.push(...estructura[tipo][subtipo]);
        }
        if (tipo === 'CAPITAL') {
            reporte.push(['  Resultado del Ejercicio', '', resultadoDelEjercicio]);
            totales.CAPITAL += resultadoDelEjercicio;
        }
        reporte.push([`TOTAL ${tipo}`, '', totales[tipo]]);
        reporte.push(['', '', '']);
    });

    reporte.push(['TOTAL PASIVO + CAPITAL', '', totales.PASIVO + totales.CAPITAL]);
    return reporte;
}

// --- Helpers de Hojas ---

function escribirReporteBalanza_(balanzaData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Balanza de Comprobación');
  sheet.clearContents().getRange(1, 1, 1, 6).setValues([['Cuenta', 'Nombre', 'S. Inicial', 'Debe', 'Haber', 'S. Final']]).setFontWeight('bold');
  if (balanzaData.length > 0) sheet.getRange(2, 1, balanzaData.length, 6).setValues(balanzaData).setNumberFormat('"$"#,##0.00');
  sheet.autoResizeColumns(1, 2);
}

function escribirReporteGeneral_(sheetName, data, columns) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    sheet.clearContents();
    sheet.getRange(1, 1, data.length, columns.length).setValues(data);
    sheet.getRange(`${columns[columns.length - 1]}:${columns[columns.length - 1]}`).setNumberFormat('"$"#,##0.00');
    sheet.getRange("A1:C1").setFontWeight('bold');
    sheet.autoResizeColumns(1, columns.length);
}

function leerHoja_(sheetName) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues(); }
function getConfiguracion_(param) { return getContabilidadConfig()[param]; }
