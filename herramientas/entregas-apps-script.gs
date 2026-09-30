/**
 * Recepción de trabajos del Portal del Alumno · Instituto Bíblico Concepción
 *
 * Guarda en Google Drive los archivos que los alumnos entregan desde la pestaña
 * "Entrega de Trabajos" del portal, y anota cada entrega en una planilla.
 * También guarda y entrega las fotos de perfil de los alumnos, y registra cada
 * ingreso al portal (hoja "Ingresos" de la misma planilla).
 *
 * Instalación (una sola vez):
 *   1. Entra a https://script.google.com con tu cuenta de Google y crea un "Proyecto nuevo".
 *   2. Borra el contenido de Código.gs y pega TODO este archivo. Guarda (Ctrl+S).
 *   3. Pulsa "Implementar" > "Nueva implementación" > tipo "Aplicación web".
 *        - Ejecutar como: Yo
 *        - Quién tiene acceso: Cualquier usuario
 *   4. Autoriza los permisos que pide (Drive y Hojas de cálculo).
 *   5. Copia la "URL de la aplicación web" (termina en /exec) y pégala en
 *      ENTREGAS_URL dentro de index.html.
 *
 * En tu Drive se crea la carpeta "Entregas Portal IBC", con una subcarpeta por
 * curso y por trabajo, la planilla "Registro de entregas" y la carpeta
 * "Fotos de alumnos" (un archivo por RUT; borra uno para quitar esa foto).
 *
 * Para actualizar el código: pega la versión nueva, guarda y luego
 * Implementar > Gestionar implementaciones > Editar (lápiz) > Versión: Nueva versión > Implementar.
 * Así la URL no cambia.
 */

const CARPETA_RAIZ = 'Entregas Portal IBC';
const PORTAL_URL = 'https://profesorrudy.github.io/index.html';
const MAX_BYTES = 10 * 1024 * 1024; // 10 MB
const CARPETA_FOTOS = 'Fotos de alumnos';
const MAX_FOTO_BYTES = 600 * 1024; // las fotos llegan ya reducidas desde el portal
const EXTENSIONES = ['doc', 'docx', 'pdf', 'odt', 'rtf', 'txt', 'ppt', 'pptx', 'jpg', 'jpeg', 'png'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.accion === 'foto') return guardarFoto(d);
    if (d.accion === 'ingreso') return registrarIngreso(d);
    const rut = String(d.rut || '').trim();
    const curso = String(d.curso || '').trim();
    const tarea = limpiar(String(d.tarea || '').trim()) || 'Trabajo';
    const nombreArchivo = String(d.archivo || '').trim();
    const ext = (nombreArchivo.split('.').pop() || '').toLowerCase();

    if (!rut || !curso || !nombreArchivo || !d.datos) return responder(false, 'Faltan datos de la entrega.');
    if (EXTENSIONES.indexOf(ext) === -1) return responder(false, 'Tipo de archivo no permitido.');

    const alumno = buscarAlumno(rut, curso);
    if (!alumno) return responder(false, 'El RUT no está matriculado en este curso.');

    const bytes = Utilities.base64Decode(d.datos);
    if (bytes.length > MAX_BYTES) return responder(false, 'El archivo supera los 10 MB.');

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const carpeta = subcarpeta(subcarpeta(raiz(), limpiar(curso)), tarea);
      const fecha = Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd HH.mm');
      const nombreFinal = curso + ' - ' + tarea + ' - ' + alumno + ' - ' + fecha + '.' + ext;
      const blob = Utilities.newBlob(bytes, d.tipo || MimeType.MICROSOFT_WORD, limpiar(nombreFinal));
      const archivo = carpeta.createFile(blob);
      const comentario = String(d.comentario || '').slice(0, 1000);
      if (comentario) archivo.setDescription(comentario);
      registro().appendRow([new Date(), alumno, rut, curso, tarea, nombreArchivo, comentario, archivo.getUrl()]);
      return responder(true, 'Trabajo recibido.', { archivo: archivo.getName() });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return responder(false, 'Error al guardar: ' + err.message);
  }
}

/* Foto de perfil: el portal la pide con ?foto=RUT al ingresar el alumno. */
function doGet(e) {
  const rut = String((e && e.parameter && e.parameter.foto) || '');
  if (!rut) return responder(false, 'Portal del Alumno · Instituto Bíblico Concepción');
  const id = PropertiesService.getScriptProperties().getProperty('foto_' + normRut(rut));
  if (!id) return responder(true, 'Sin foto.', { foto: null });
  try {
    const f = DriveApp.getFileById(id);
    if (f.isTrashed()) return responder(true, 'Sin foto.', { foto: null });
    const blob = f.getBlob();
    return responder(true, 'Foto.', { foto: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) });
  } catch (err) {
    return responder(true, 'Sin foto.', { foto: null });
  }
}

function guardarFoto(d) {
  const rut = String(d.rut || '').trim();
  if (!rut || !d.datos) return responder(false, 'Faltan datos de la foto.');
  const alumno = buscarAlumno(rut, null);
  if (!alumno) return responder(false, 'RUT no habilitado en el portal.');
  const bytes = Utilities.base64Decode(d.datos);
  if (bytes.length > MAX_FOTO_BYTES) return responder(false, 'La foto es demasiado grande.');
  const props = PropertiesService.getScriptProperties();
  const clave = 'foto_' + normRut(rut);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const anterior = props.getProperty(clave);
    if (anterior) { try { DriveApp.getFileById(anterior).setTrashed(true); } catch (e) { /* ya no existe */ } }
    const carpeta = subcarpeta(raiz(), CARPETA_FOTOS);
    const archivo = carpeta.createFile(Utilities.newBlob(bytes, 'image/jpeg', limpiar(alumno + ' - ' + rut) + '.jpg'));
    props.setProperty(clave, archivo.getId());
    return responder(true, 'Foto actualizada.');
  } finally {
    lock.releaseLock();
  }
}

/* Guarda la fecha de este ingreso y devuelve la del ingreso anterior. */
function registrarIngreso(d) {
  const rut = String(d.rut || '').trim();
  if (!rut) return responder(false, 'Falta el RUT.');
  const alumno = buscarAlumno(rut, null);
  if (!alumno) return responder(false, 'RUT no habilitado en el portal.');
  const props = PropertiesService.getScriptProperties();
  const clave = 'ingreso_' + normRut(rut);
  const ahora = new Date();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const anterior = props.getProperty(clave);
    props.setProperty(clave, ahora.toISOString());
    let hoja = registro().getParent().getSheetByName('Ingresos');
    if (!hoja) {
      hoja = registro().getParent().insertSheet('Ingresos');
      hoja.appendRow(['Fecha', 'Alumno', 'RUT']);
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([ahora, alumno, rut]);
    return responder(true, 'Ingreso registrado.', { anterior: anterior });
  } finally {
    lock.releaseLock();
  }
}

function normRut(r) { return String(r).replace(/[\s.\-]/g, '').toUpperCase(); }

/* Comprueba en el portal publicado que el RUT exista y no esté bloqueado;
   si se indica un curso, además que esté matriculado en él. */
function buscarAlumno(rut, curso) {
  const cache = CacheService.getScriptCache();
  let html = cache.get('portal');
  if (!html) {
    html = UrlFetchApp.fetch(PORTAL_URL, { muteHttpExceptions: true }).getContentText();
    try { cache.put('portal', html, 600); } catch (e) { /* el portal es muy grande para la caché */ }
  }
  const lineas = html.split('\n');
  for (const l of lineas) {
    const m = l.match(/^\s*"([0-9.]+-[0-9Kk])": \{ nombre: "([^"]+)"/);
    if (m && normRut(m[1]) === normRut(rut)) {
      if (l.indexOf('bloqueado: true') !== -1) return null;
      if (curso && l.indexOf('"' + curso + '"') === -1) return null;
      return m[2];
    }
  }
  return null;
}

function raiz() {
  const it = DriveApp.getFoldersByName(CARPETA_RAIZ);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_RAIZ);
}

function subcarpeta(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

function registro() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('registroId');
  if (id) {
    try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) { /* se borró: se crea otra */ }
  }
  const ss = SpreadsheetApp.create('Registro de entregas');
  DriveApp.getFileById(ss.getId()).moveTo(raiz());
  const hoja = ss.getSheets()[0];
  hoja.appendRow(['Fecha', 'Alumno', 'RUT', 'Curso', 'Trabajo', 'Archivo original', 'Comentario', 'Enlace']);
  hoja.setFrozenRows(1);
  props.setProperty('registroId', ss.getId());
  return hoja;
}

function limpiar(t) { return t.replace(/[\\/:*?"<>|]/g, '-').slice(0, 180); }

function responder(ok, mensaje, extra) {
  return ContentService.createTextOutput(JSON.stringify(Object.assign({ ok: ok, mensaje: mensaje }, extra || {})))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Prueba rápida desde el editor: Ejecutar > probar. Crea la carpeta y la planilla. */
function probar() {
  raiz();
  registro();
  Logger.log('Carpeta y planilla listas en tu Drive.');
}
