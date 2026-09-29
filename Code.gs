// Configuración del proyecto
const CONFIG = {
  SPREADSHEET_ID: '1JnoBYqSlamwNx23kV8y5thneDa4YMCyMZMvgIn5RAZo',
  DRIVE_FOLDER_ID: '17Sc-b2A0D5YwsCEcvb1dA6cT48tPaf_i',
  SHEET_NAME: 'Incidencias'
};

function validarConfiguracion() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID.includes('PEGA_AQUI')) {
    throw new Error('Configura CONFIG.SPREADSHEET_ID con el ID real de tu hoja de Google Sheets.');
  }

  if (!CONFIG.DRIVE_FOLDER_ID || CONFIG.DRIVE_FOLDER_ID.includes('PEGA_AQUI')) {
    throw new Error('Configura CONFIG.DRIVE_FOLDER_ID con el ID real de la carpeta de Google Drive.');
  }
}

/**
 * Renderiza la interfaz principal
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Control de Incidencias en Planteles')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Obtiene todos los registros para la tabla/dashboard
 */
function getIncidencias() {
  try {
    console.log('Paso 1: validando configuración');
    validarConfiguracion();

    console.log('Paso 2: abriendo hoja de cálculo');
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);

    console.log('Paso 3: leyendo datos');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];

    const rows = data.slice(1);
    
    console.log('Paso 4: mapeando registros');
    return rows.map(row => ({
      id: row[0],
      fecha: row[1],
      plantel: row[2],
      tipo: row[3],
      prioridad: row[4],
      descripcion: row[5],
      estatus: row[6],
      fotoInicial: row[7],
      fotoSolucion: row[8],
      comentarios: row[9],
      usuario: row[10]
    })).reverse();
  } catch (error) {
    console.error('FALLÓ en getIncidencias:', error);
    throw new Error('Error al leer datos en paso: ' + error.message);
  }
}

/**
 * Guarda un nuevo reporte enviado desde el formulario
 */
function registrarIncidencia(form) {
  try {
    console.log('Paso 1: validando configuración');
    validarConfiguracion();

    console.log('Paso 2: abriendo hoja de cálculo');
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);

    console.log('Paso 3: generando ID y datos');
    const id = 'INC-' + new Date().getTime().toString().slice(-6);
    const fecha = new Date().toLocaleString('es-MX');
    const usuario = Session.getActiveUser().getEmail() || 'Operador Local';

    let fotoUrl = '';
    if (form.fotoBase64) {
      try {
        console.log('Paso 4: guardando foto inicial en Drive');
        fotoUrl = guardarEnDrive(form.fotoBase64, `${id}_inicio.jpg`);
      } catch (error) {
        console.warn('No se pudo guardar la foto inicial en Drive:', error.message);
        fotoUrl = '';
      }
    }

    console.log('Paso 5: escribiendo fila en la hoja');
    sheet.appendRow([
      id,
      fecha,
      form.plantel,
      form.tipo,
      form.prioridad,
      form.descripcion,
      'Nuevo', // Estatus inicial
      fotoUrl,
      '', // Foto solución vacía
      '', // Sin comentarios de inicio
      usuario
    ]);

    console.log('Paso 6: registro terminado correctamente');
    return { success: true, id: id };
  } catch (error) {
    console.error('FALLÓ en registrarIncidencia:', error);
    return { success: false, error: 'Error en registrarIncidencia: ' + error.toString() };
  }
}

/**
 * Actualiza el estado de la incidencia y guarda evidencia de solución
 */
function actualizarEstatusIncidencia(payload) {
  try {
    console.log('Paso 1: validando configuración');
    validarConfiguracion();

    console.log('Paso 2: abriendo hoja y leyendo datos');
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let filaIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        filaIndex = i + 1;
        break;
      }
    }

    if (filaIndex === -1) throw new Error('Registro no encontrado.');

    let fotoSolucionUrl = data[filaIndex - 1][8];
    if (payload.fotoBase64) {
      try {
        console.log('Paso 3: guardando foto de solución en Drive');
        fotoSolucionUrl = guardarEnDrive(payload.fotoBase64, `${payload.id}_solucion.jpg`);
      } catch (error) {
        console.warn('No se pudo guardar la foto de solución en Drive:', error.message);
        fotoSolucionUrl = data[filaIndex - 1][8];
      }
    }

    console.log('Paso 4: actualizando estatus y comentarios');
    sheet.getRange(filaIndex, 7).setValue(payload.nuevoEstatus);
    sheet.getRange(filaIndex, 9).setValue(fotoSolucionUrl);
    
    const comentariosPrevios = data[filaIndex - 1][9] || '';
    const fechaHoy = new Date().toLocaleDateString('es-MX');
    const nuevoHistorial = comentariosPrevios 
      ? `${comentariosPrevios}\n[${fechaHoy}]: ${payload.comentario}` 
      : `[${fechaHoy}]: ${payload.comentario}`;
      
    sheet.getRange(filaIndex, 10).setValue(nuevoHistorial);

    console.log('Paso 5: actualización finalizada');
    return { success: true };
  } catch (error) {
    console.error('FALLÓ en actualizarEstatusIncidencia:', error);
    return { success: false, error: 'Error en actualizarEstatusIncidencia: ' + error.toString() };
  }
}

/**
 * Función auxiliar para guardar imágenes en Base64 a Google Drive
 */
function guardarEnDrive(base64Data, fileName) {
  if (!base64Data || !base64Data.includes('base64,')) {
    console.log('guardarEnDrive: no hay base64 para guardar');
    return '';
  }

  try {
    console.log('guardarEnDrive: validando configuración');
    validarConfiguracion();

    console.log('guardarEnDrive: abriendo carpeta de Drive');
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

    console.log('guardarEnDrive: decodificando imagen');
    const parts = base64Data.split(';base64,');
    const contentType = parts[0].replace('data:', '');
    const decodedData = Utilities.base64Decode(parts[1]);
    const blob = Utilities.newBlob(decodedData, contentType, fileName);

    console.log('guardarEnDrive: creando archivo');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    console.log('guardarEnDrive: archivo creado correctamente');
    return file.getUrl();
  } catch (error) {
    console.error('FALLÓ en guardarEnDrive:', error);
    throw new Error('Acceso a Drive denegado o carpeta no válida: ' + error.message);
  }
}