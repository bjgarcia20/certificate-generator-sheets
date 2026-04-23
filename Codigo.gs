// ===== CONFIGURACIÓN =====
// Los IDs se guardan en las Propiedades del Script (no en este código).
// Ver instrucciones abajo para configurarlos.

const COLUMNA_PLANTILLA = "Plantilla";
const COLUMNA_NOMBRE_ESTUDIANTE = "nombre_estudiante";

let cachePlantillas = {};

// ===== LECTURA DE IDs =====
function getCarpetaPlantillasId() {
  const id = PropertiesService.getScriptProperties().getProperty("CARPETA_PLANTILLAS_ID");
  if (!id) throw new Error("Falta configurar CARPETA_PLANTILLAS_ID en Propiedades del Script.");
  return id;
}

function getCarpetaDestinoId() {
  const id = PropertiesService.getScriptProperties().getProperty("CARPETA_DESTINO_ID");
  if (!id) throw new Error("Falta configurar CARPETA_DESTINO_ID en Propiedades del Script.");
  return id;
}

// ===== MENÚ =====
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📄 Documentos")
    .addItem("Generar certificados", "generarCertificados")
    .addItem("Certificado (fila actual)", "generarCertificadoFila")
    .addToUi();
}

// ===== GENERACIÓN MASIVA =====
function generarCertificados() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];
  const carpetaDestino = DriveApp.getFolderById(getCarpetaDestinoId());

  cachePlantillas = {};
  let exitosos = 0;
  let errores = [];

  for (let i = 1; i < datos.length; i++) {
    let fila = datos[i];
    if (fila[0] == "") continue;

    try {
      crearCertificado(fila, encabezados, carpetaDestino);
      exitosos++;
    } catch (e) {
      errores.push(`Fila ${i + 1}: ${e.message}`);
    }
  }

  let mensaje = `✅ ${exitosos} certificado(s) generado(s).`;
  if (errores.length > 0) {
    mensaje += `\n\n⚠️ Errores:\n${errores.join("\n")}`;
  }
  SpreadsheetApp.getUi().alert(mensaje);
}

// ===== GENERACIÓN FILA ACTUAL =====
function generarCertificadoFila() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const filaIndex = hoja.getActiveRange().getRow();

  if (filaIndex == 1) {
    SpreadsheetApp.getUi().alert("Selecciona una fila válida");
    return;
  }

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];
  const fila = datos[filaIndex - 1];
  const carpetaDestino = DriveApp.getFolderById(getCarpetaDestinoId());

  cachePlantillas = {};

  try {
    crearCertificado(fila, encabezados, carpetaDestino);
    SpreadsheetApp.getUi().alert("✅ Certificado generado correctamente.");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`⚠️ Error: ${e.message}`);
  }
}

// ===== LÓGICA PRINCIPAL =====
function crearCertificado(fila, encabezados, carpetaDestino) {
  const indicePlantilla = encabezados.indexOf(COLUMNA_PLANTILLA);
  if (indicePlantilla === -1) {
    throw new Error(`No se encontró la columna "${COLUMNA_PLANTILLA}".`);
  }

  const indiceNombre = encabezados.indexOf(COLUMNA_NOMBRE_ESTUDIANTE);
  if (indiceNombre === -1) {
    throw new Error(`No se encontró la columna "${COLUMNA_NOMBRE_ESTUDIANTE}".`);
  }

  const nombrePlantilla = fila[indicePlantilla];
  const nombreEstudiante = fila[indiceNombre];

  if (!nombrePlantilla) throw new Error("La fila no tiene plantilla asignada.");
  if (!nombreEstudiante) throw new Error("La fila no tiene nombre de estudiante.");

  const plantilla = obtenerPlantilla(nombrePlantilla);
  const nombrePlantillaLimpio = limpiarNombrePlantilla(plantilla.getName());
  const nombreArchivo = `${nombreEstudiante} - ${nombrePlantillaLimpio}`;

  const copia = plantilla.makeCopy(nombreArchivo, carpetaDestino);
  const doc = DocumentApp.openById(copia.getId());
  const body = doc.getBody();

  for (let j = 0; j < encabezados.length; j++) {
    let campo = `{{${encabezados[j]}}}`;
    let valor = fila[j] != null ? fila[j].toString() : "";

    if (j === indiceNombre) {
      valor = valor.toUpperCase();
    }

    body.replaceText(campo, valor);
  }

  doc.saveAndClose();
}

// ===== UTILIDADES =====
function obtenerPlantilla(nombreBuscado) {
  const clave = normalizar(nombreBuscado);

  if (cachePlantillas[clave]) {
    return cachePlantillas[clave];
  }

  const carpetaPlantillas = DriveApp.getFolderById(getCarpetaPlantillasId());
  const archivos = carpetaPlantillas.getFiles();

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (normalizar(archivo.getName()) === clave) {
      cachePlantillas[clave] = archivo;
      return archivo;
    }
  }

  throw new Error(`No se encontró la plantilla "${nombreBuscado}" en la carpeta.`);
}

function limpiarNombrePlantilla(nombre) {
  return nombre.replace(/^\d+\.\s*PLANTILLA\s+/i, "").trim();
}

function normalizar(texto) {
  return texto
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}