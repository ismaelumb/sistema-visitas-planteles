// Configuración del proyecto
const CONFIG = {
  SPREADSHEET_ID: 'PEGA_AQUI_TU_ID_DE_GOOGLE_SHEETS',
  DRIVE_FOLDER_ID: 'PEGA_AQUI_TU_ID_DE_CARPETA_DRIVE',
  SHEET_NAME: 'Incidencias'
};

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
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];

    const rows = data.slice(1);
    
    // Mapeo simple a objeto JSON
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
    })).reverse(); // Los más recientes primero
  } catch (error) {
    throw new Error('Error al leer datos: ' + error.message);
  }
}

/**
 * Guarda un nuevo reporte enviado desde el formulario
 */
function registrarIncidencia(form) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const id = 'INC-' + new Date().getTime().toString().slice(-6);
    const fecha = new Date().toLocaleString('es-MX');
    const usuario = Session.getActiveUser().getEmail() || 'Operador Local';

    let fotoUrl = '';
    if (form.fotoBase64) {
      fotoUrl = guardarEnDrive(form.fotoBase64, `${id}_inicio.jpg`);
    }

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

    return { success: true, id: id };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Actualiza el estado de la incidencia y guarda evidencia de solución
 */
function actualizarEstatusIncidencia(payload) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    let filaIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.id) {
        filaIndex = i + 1; // Índice en Apps Script inicia en 1
        break;
      }
    }

    if (filaIndex === -1) throw new Error('Registro no encontrado.');

    let fotoSolucionUrl = data[filaIndex - 1][8];
    if (payload.fotoBase64) {
      fotoSolucionUrl = guardarEnDrive(payload.fotoBase64, `${payload.id}_solucion.jpg`);
    }

    // Actualizar campos en el Sheet
    sheet.getRange(filaIndex, 7).setValue(payload.nuevoEstatus); // Columna G: Estatus
    sheet.getRange(filaIndex, 9).setValue(fotoSolucionUrl);       // Columna I: Foto Solución
    
    // Concatenar historial de comentarios
    const comentariosPrevios = data[filaIndex - 1][9] || '';
    const fechaHoy = new Date().toLocaleDateString('es-MX');
    const nuevoHistorial = comentariosPrevios 
      ? `${comentariosPrevios}\n[${fechaHoy}]: ${payload.comentario}` 
      : `[${fechaHoy}]: ${payload.comentario}`;
      
    sheet.getRange(filaIndex, 10).setValue(nuevoHistorial);      // Columna J: Comentarios

    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Función auxiliar para guardar imágenes en Base64 a Google Drive
 */
function guardarEnDrive(base64Data, fileName) {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const parts = base64Data.split(';base64,');
  const contentType = parts[0].replace('data:', '');
  const decodedData = Utilities.base64Decode(parts[1]);
  const blob = Utilities.newBlob(decodedData, contentType, fileName);
  
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}