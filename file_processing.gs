/**
 * @file file_processing.gs
 * @description Módulo para manejar la carga y el procesamiento de archivos XML de CFDI.
 *
 * @author Jules
 * @version 1.1.0
 */

/**
 * Procesa los archivos XML cargados por el usuario.
 * @param {Array<Object>} files - Array de objetos { name: string, content: string }.
 * @returns {Object} Resultado del procesamiento { status: 'success' | 'error', message: string }.
 */
function procesarContenidoXMLs(files) {
  if (!files || files.length === 0) {
    return { status: 'error', message: 'No se recibieron archivos.' };
  }

  try {
    const config = getContabilidadConfig();
    const periodoActual = config['Periodo Contable Actual (YYYY-MM)'];
    const nuestroRfc = config['RFC de la Empresa']; // Asume que este campo existe en Configuración

    if (!periodoActual || !nuestroRfc) {
      throw new Error('"Periodo Contable Actual" o "RFC de la Empresa" no están definidos en la hoja "Configuración".');
    }

    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Log_Procesados');
    const processedUuids = new Set(logSheet.getRange(2, 1, logSheet.getLastRow(), 1).getValues().flat());

    let successCount = 0;
    let errorCount = 0;
    let skippedCount = 0;

    files.forEach(file => {
      try {
        const timbre = getTimbreFiscal(file.content);
        if (!timbre.uuid) {
          Logger.log(`Archivo ${file.name} sin Timbre Fiscal válido.`);
          errorCount++;
          return;
        }

        if (processedUuids.has(timbre.uuid)) {
          Logger.log(`UUID ${timbre.uuid} ya procesado.`);
          skippedCount++;
          return;
        }

        const cfdiData = parsearCFDI(file.content, timbre);
        if (!cfdiData) {
          Logger.log(`Error al parsear o validar ${file.name} (UUID: ${timbre.uuid}).`);
          errorCount++;
          return;
        }

        // --- Validación del Periodo Contable ---
        const cfdiPeriodo = cfdiData.fecha.substring(0, 7); // Formato YYYY-MM
        if (cfdiPeriodo !== periodoActual) {
          Logger.log(`CFDI ${timbre.uuid} (${cfdiPeriodo}) no pertenece al periodo actual (${periodoActual}). Se omite.`);
          skippedCount++;
          return;
        }

        const poliza = generarPoliza(cfdiData, nuestroRfc);
        if (poliza && poliza.length > 0) {
          registrarPolizaEnSheet(poliza);
          registrarEnLog(cfdiData.uuid);
          processedUuids.add(cfdiData.uuid);
          successCount++;
        } else {
          Logger.log(`No se generó póliza para ${timbre.uuid}.`);
          errorCount++;
        }
      } catch (e) {
        Logger.log(`Error procesando ${file.name}: ${e.message}`);
        errorCount++;
      }
    });

    const message = `Proceso finalizado. Éxitos: ${successCount}, Errores: ${errorCount}, Omitidos: ${skippedCount}.`;
    return { status: 'success', message: message };

  } catch (e) {
    Logger.log(`Error crítico en procesarContenidoXMLs: ${e.message}`);
    return { status: 'error', message: `Error crítico: ${e.message}` };
  }
}
