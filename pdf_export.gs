/**
 * Orquesta la exportación del reporte consolidado a PDF usando una hoja temporal.
 */
function exportReportAsPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tempSheetName = 'ReportePDF_Temp';
  let tempSheet = ss.getSheetByName(tempSheetName);
  if (tempSheet) {
    ss.deleteSheet(tempSheet);
  }
  tempSheet = ss.insertSheet(tempSheetName, 0);

  try {
    ui.showToast('Iniciando exportación...', 'Preparando Reporte', -1);

    // 1. Ensamblar el contenido en la hoja temporal
    let currentRow = 1;
    currentRow = createCoverPage(tempSheet, currentRow);
    currentRow = copyReportSection(tempSheet, currentRow, 'Inicio', 'K1:R20', 'Reportes Financieros');

    // Placeholders para secciones futuras
    tempSheet.getRange(currentRow, 1).setValue('Análisis de Impuestos (IVA/ISR)').setFontSize(14).setFontWeight('bold');
    tempSheet.getRange(currentRow + 1, 1).setValue('[Funcionalidad en desarrollo]');
    currentRow += 3;

    tempSheet.getRange(currentRow, 1).setValue('Nota del Contador').setFontSize(14).setFontWeight('bold');
    tempSheet.getRange(currentRow + 1, 1).setValue('[Espacio para notas y recomendaciones]');

    // 2. Exportar la hoja temporal a PDF
    const allSheets = ss.getSheets();
    allSheets.forEach(sheet => {
      if (sheet.getName() !== tempSheetName) {
        sheet.hideSheet();
      }
    });

    SpreadsheetApp.flush(); // Aplicar cambios

    const pdfBlob = ss.getBlob().getAs('application/pdf');
    const contribuyente = "Contribuyente"; // Placeholder
    const periodo = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM yyyy");
    pdfBlob.setName(`Reporte Contable - ${contribuyente} - ${periodo}.pdf`);

    const pdfFile = DriveApp.createFile(pdfBlob);
    ui.showToast('Exportación Exitosa!', `Archivo guardado en Drive: ${pdfFile.getName()}`, 5);

  } catch (e) {
    Logger.log(e);
    ui.alert('Error', `Ocurrió un error al exportar: ${e.message}`, ui.ButtonSet.OK);
  } finally {
    // 3. Limpieza: borrar hoja temporal y mostrar las originales
    const originalSheets = ss.getSheets();
    originalSheets.forEach(sheet => {
      if(sheet.getName() !== tempSheetName) {
        sheet.showSheet();
      }
    });
    if (ss.getSheetByName(tempSheetName)) {
      ss.deleteSheet(ss.getSheetByName(tempSheetName));
    }
    ui.showToast('Proceso finalizado.');
  }
}

function createCoverPage(sheet, startRow) {
  sheet.getRange(startRow, 1, 1, 5).merge()
    .setValue('Reporte Financiero Contable')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground(THEME_COLORS.BLUE_ACCENT)
    .setRowHeight(50);

  sheet.getRange(startRow + 2, 1).setValue('Contribuyente:');
  sheet.getRange(startRow + 2, 2).setValue('Mi Empresa S.A. de C.V.'); // Placeholder
  sheet.getRange(startRow + 3, 1).setValue('Período:');
  sheet.getRange(startRow + 3, 2).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM yyyy"));

  return startRow + 5;
}

function copyReportSection(targetSheet, startRow, sourceSheetName, sourceRange, title) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  if (!sourceSheet) return startRow;

  targetSheet.getRange(startRow, 1).setValue(title).setFontSize(14).setFontWeight('bold');
  startRow++;

  const rangeToCopy = sourceSheet.getRange(sourceRange);
  rangeToCopy.copyTo(targetSheet.getRange(startRow, 1), {contentsOnly: false, formatOnly: false});

  return startRow + rangeToCopy.getNumRows() + 2; // Añadir espacio
}
