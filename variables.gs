function abrirHojaPorID(id) {
  return SpreadsheetApp.openById(id).getSheets()[0];
}

function leerDatos(hoja) {
  const filas = hoja.getLastRow();
  const columnas = hoja.getLastColumn();
  // Si solo hay encabezados o está vacía, salimos
  if (filas <= 1) return []; 
  return hoja.getRange(2, 1, filas - 1, columnas).getValues();
}

function obtenerNombre(filaActual) { return filaActual[0]; }
function obtenerCurso(filaActual) { return filaActual[1]; }

function obtenerFecha(filaActual) {
  // Manejo básico de errores: si no hay fecha, enviamos un texto vacío
  if (!filaActual[2]) return "Sin fecha";
  let fecha = new Date(filaActual[2]);
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function crearCopiaPlantilla(id) {
  return DriveApp.getFileById(id).makeCopy();
}

function abrirDocumento(archivoCopia) {
  return DocumentApp.openById(archivoCopia.getId());
}

function reemplazarTexto(doc, etiqueta, valor) {
  const cuerpo = doc.getBody();
  // Aseguramos que el valor sea texto con String() para evitar errores si viene un número
  cuerpo.replaceText(etiqueta, String(valor));
}

function guardarDocumento(doc) {
  doc.saveAndClose();
}

function cambiarNombreDocumento(archivo, nombre) {
  archivo.setName("Constancia - " + nombre);
}

function moverArchivoACarpeta(archivo, idCarpeta) {
  const carpeta = DriveApp.getFolderById(idCarpeta);
  carpeta.addFile(archivo);
  
 
}
