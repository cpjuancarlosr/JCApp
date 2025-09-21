/**
 * @file sat_connector.gs
 * @description Módulo para la conexión y descarga de CFDI desde el portal del SAT.
 * Incluye la lógica para manejar la autenticación con FIEL y el procesamiento de las descargas.
 * Contiene una ruta alterna para procesar XML desde una carpeta de Drive si la conexión directa falla.
 *
 * @author Jules
 * @version 1.0.0
 */

/**
 * Muestra una interfaz de usuario HTML para que el usuario configure las rutas
 * a sus archivos .cer y .key de la FIEL y su contraseña.
 */
function showFielConfigurationUi() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('FielConfigUi')
      .setWidth(400)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Configuración de FIEL');
}

/**
 * Guarda la configuración de la FIEL. Es llamada desde la UI de HTML.
 * @param {object} config - Objeto con certPath, keyPath y password.
 */
function saveFielConfig(config) {
  try {
    // Guardar rutas en la hoja CFG
    setConfigValue('Ruta Archivo .CER en Drive', config.certPath);
    setConfigValue('Ruta Archivo .KEY en Drive', config.keyPath);

    // Guardar contraseña de forma segura en las propiedades del usuario para la sesión actual
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('fielPassword', config.password);

    return { status: 'success', message: 'Configuración guardada exitosamente. La contraseña es válida solo para esta sesión.' };
  } catch (e) {
    return { status: 'error', message: `Error al guardar: ${e.message}` };
  }
}

/**
 * Función principal para orquestar la descarga de CFDIs.
 * Intenta la conexión directa y, si falla, instruye al usuario sobre cómo
 * usar la carpeta de Drive como alternativa.
 */
function downloadCfdis() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Descarga de CFDI', 'Ingrese el período (formato AAAA-MM) y tipo (Emitidos/Recibidos). Ej: 2023-12, Recibidos', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const [period, type] = response.getResponseText().split(',').map(s => s.trim());

    // Lógica de intento de descarga directa (placeholder)
    const success = attemptDirectDownload(period, type);

    if (!success) {
      ui.alert(
        'La descarga directa no está disponible o falló.',
        'Por favor, suba sus archivos XML manualmente a la carpeta de Drive configurada en CFG y luego ejecute la opción "Procesar CFDI desde Carpeta".',
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * Intenta realizar la descarga directa (simulación).
 * @returns {boolean} `false` para simular que la conexión directa no está implementada.
 */
function attemptDirectDownload(period, type) {
  Logger.log(`Intento de descarga para ${period}, tipo ${type}. Esta función es un placeholder.`);
  // Aquí iría la compleja lógica de autenticación y descarga SOAP.
  // Por ahora, siempre retornará false para activar la ruta alterna.
  return false;
}

/**
 * Procesa los archivos XML que se encuentran en la carpeta de Drive especificada en la hoja CFG.
 * Mueve los archivos procesados a una subcarpeta "Archivados".
 */
function processCfdisFromDriveFolder() {
  const ui = SpreadsheetApp.getUi();
  const folderId = getConfigValue('ID Carpeta XML en Drive');

  if (!folderId) {
    ui.alert('Configuración Requerida', 'Por favor, ingrese el ID de la carpeta de Google Drive que contiene los XML en la hoja "CFG", celda B1.', ui.ButtonSet.OK);
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);

    // Crear o encontrar la carpeta de archivos procesados
    let archiveFolder;
    const archiveFolders = folder.getFoldersByName('Archivados');
    if (archiveFolders.hasNext()) {
      archiveFolder = archiveFolders.next();
    } else {
      archiveFolder = folder.createFolder('Archivados');
    }

    const files = folder.getFilesByType(MimeType.XML);
    const allParsedData = [];
    let processedCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const content = file.getBlob().getDataAsString('UTF-8');
      const parsedData = parseXml(content); // Llama al parser en xml_parser.gs

      if (parsedData) {
        allParsedData.push(parsedData);
        // Mover archivo a la carpeta de archivados para no procesarlo de nuevo
        file.moveTo(archiveFolder);
        processedCount++;
      } else {
        Logger.log(`No se pudo parsear el archivo: ${file.getName()}`);
      }
    }

    if (allParsedData.length > 0) {
      updateXmlSheet(allParsedData); // Llama a una función en xml_parser.gs
    } else {
      ui.alert('Proceso Terminado', 'No se encontraron nuevos archivos XML para procesar.', ui.ButtonSet.OK);
    }

  } catch (e) {
    Logger.log(e);
    ui.alert('Error', `Ocurrió un error al procesar los archivos: ${e.message}. Verifique que el ID de la carpeta sea correcto y que tenga permisos de acceso.`, ui.ButtonSet.OK);
  }
}
