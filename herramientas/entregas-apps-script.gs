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
 * Panel del profesor: en Configuración del proyecto (engranaje) > Propiedades de
 * la secuencia de comandos, agrega la propiedad CLAVE_PROFESOR con tu clave.
 * El panel del portal pide esa clave para mostrar ingresos y entregas.
 *
 * Planilla de datos (opcional): sube "Portal IBC - Datos" a tu Drive como Hojas de
 * cálculo de Google y agrega la propiedad PLANILLA_DATOS con su ID (lo que va entre
 * /d/ y /edit en su dirección). Desde ese momento el portal lee de ahí alumnos,
 * notas, asistencia, materiales, tareas, horarios, observaciones y avisos.
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
    if (d.accion === 'panel') return datosPanel(d);
    if (d.accion === 'datos') return datosAlumno(d);
    const rut = String(d.rut || '').trim();
    const curso = String(d.curso || '').trim();
    const tarea = limpiar(String(d.tarea || '').trim()) || 'Trabajo';
    const nombreArchivo = String(d.archivo || '').trim();
    const ext = (nombreArchivo.split('.').pop() || '').toLowerCase();

    if (!rut || !curso || !nombreArchivo || !d.datos) return responder(false, 'Faltan datos de la entrega.');
    if (EXTENSIONES.indexOf(ext) === -1) return responder(false, 'Tipo de archivo no permitido.');

    const alumno = buscarAlumno(rut, curso, d.clave);
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
  const alumno = buscarAlumno(rut, null, d.clave);
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
  const alumno = buscarAlumno(rut, null, d.clave);
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

/* Panel del profesor: devuelve los registros de entregas e ingresos si la clave es correcta. */
function datosPanel(d) {
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE_PROFESOR');
  if (!clave) return responder(false, 'Falta configurar CLAVE_PROFESOR en las propiedades del script.');
  if (String(d.clave || '') !== clave) {
    Utilities.sleep(1500); // frena intentos de adivinar la clave
    return responder(false, 'Clave incorrecta.');
  }
  const libro = registro().getParent();
  const filas = hoja => {
    if (!hoja || hoja.getLastRow() < 2) return [];
    return hoja.getRange(2, 1, hoja.getLastRow() - 1, hoja.getLastColumn()).getValues();
  };
  const iso = v => (v instanceof Date ? v.toISOString() : String(v));
  const entregas = filas(libro.getSheets()[0]).map(f => ({
    fecha: iso(f[0]), alumno: f[1], rut: f[2], curso: f[3], trabajo: f[4], archivo: f[5], comentario: f[6], enlace: f[7]
  }));
  const ingresos = filas(libro.getSheetByName('Ingresos')).map(f => ({ fecha: iso(f[0]), alumno: f[1], rut: f[2] }));
  const extra = { entregas: entregas, ingresos: ingresos, planilla: libro.getUrl() };
  const datos = leerDatos();
  if (datos) {
    // Con planilla de datos, el portal ya no trae alumnos ni asistencia en su código.
    const alumnos = {}, asistencia = {};
    for (const a of Object.values(datos.alumnos)) {
      alumnos[a.rut] = { nombre: a.nombre, bloqueado: a.bloqueado, cursos: a.cursos };
      asistencia[a.rut] = datos.asistencia[normRut(a.rut)] || {};
    }
    extra.datos = { alumnos: alumnos, asistencia: asistencia, tareas: datos.tareas, cursosCursados: datos.cursosCursados };
    extra.datosPlanilla = datos.url;
  }
  return responder(true, 'Datos del panel.', extra);
}

function normRut(r) { return String(r).replace(/[\s.\-]/g, '').toUpperCase(); }

/* ===================== PLANILLA DE DATOS ===================== */

/* Ingreso de un alumno: valida RUT (y clave, si el alumno tiene una) y devuelve solo sus datos. */
function datosAlumno(d) {
  const datos = leerDatos();
  if (!datos) return responder(false, 'La planilla de datos no está configurada.', { codigo: 'sin_planilla' });
  const a = datos.alumnos[normRut(d.rut || '')];
  if (!a) return responder(false, 'RUT no encontrado. Verifica que esté bien escrito, incluyendo el guión y el dígito verificador.', { codigo: 'rut' });
  if (a.bloqueado) return responder(false, 'Tu acceso al portal está temporalmente restringido. Contacta a tu profesor.', { codigo: 'bloqueado' });
  if (a.clave) {
    if (!d.clave) return responder(false, 'Ingresa tu clave.', { codigo: 'clave' });
    if (String(d.clave).trim() !== a.clave) {
      Utilities.sleep(1500); // frena intentos de adivinar claves
      return responder(false, 'Clave incorrecta. Si no la recuerdas, pídela a tu profesor.', { codigo: 'clave_mala' });
    }
  }
  const cursos = Object.keys(a.cursos);
  const deCursos = obj => { const o = {}; for (const c of cursos) if (obj[c]) o[c] = obj[c]; return o; };
  return responder(true, 'Datos del alumno.', {
    alumno: { rut: a.rut, nombre: a.nombre, cursos: a.cursos },
    cursadas: a.cursadas,
    asistencia: datos.asistencia[normRut(a.rut)] || {},
    materiales: deCursos(datos.materiales),
    tareas: deCursos(datos.tareas),
    trabajoGrupo: deCursos(datos.trabajoGrupo),
    horarios: deCursos(datos.horarios),
    observacionesCurso: deCursos(datos.observacionesCurso),
    observaciones: datos.observaciones[normRut(a.rut)] || [],
    avisos: datos.avisos.filter(x => !x.curso || cursos.indexOf(x.curso) !== -1)
  });
}

/* Lee toda la planilla y la convierte al formato del portal (en caché 30 s). */
function leerDatos() {
  const id = PropertiesService.getScriptProperties().getProperty('PLANILLA_DATOS');
  if (!id) return null;
  const cache = CacheService.getScriptCache();
  const enCache = cache.get('datos');
  if (enCache) return JSON.parse(enCache);

  const libro = SpreadsheetApp.openById(id.trim());
  const tabla = nombre => {
    const h = libro.getSheetByName(nombre);
    if (!h || h.getLastRow() < 2) return [];
    return h.getRange(2, 1, h.getLastRow() - 1, h.getLastColumn()).getDisplayValues()
      .map(f => f.map(v => String(v).trim()))
      .filter(f => f.some(v => v !== ''));
  };
  const porCurso = (filas, armar) => {
    const o = {};
    for (const f of filas) { if (!f[0]) continue; (o[f[0]] = o[f[0]] || []).push(armar(f)); }
    return o;
  };

  const datos = { url: libro.getUrl(), alumnos: {}, asistencia: {}, observaciones: {}, cursosCursados: [] };
  for (const f of tabla('Alumnos')) {
    if (!f[0]) continue;
    datos.alumnos[normRut(f[0])] = { rut: f[0], nombre: f[1], clave: f[2], bloqueado: /^s[ií]/i.test(f[3]), cursos: {}, cursadas: [] };
  }
  const cursando = {}, cursada = {};
  for (const f of tabla('Notas')) {
    const a = datos.alumnos[normRut(f[0])];
    if (!a || !f[2]) continue;
    a.cursos[f[2]] = { t1: nota(f[3]), t2: nota(f[4]), ex: nota(f[5]), notaFinal: nota(f[6]) };
    if (/cursada/i.test(f[7])) { a.cursadas.push(f[2]); cursada[f[2]] = true; } else cursando[f[2]] = true;
  }
  datos.cursosCursados = Object.keys(cursada).filter(c => !cursando[c]);

  const estados = { '/': 'Presente', 'x': 'Ausente', '%': 'Media clase', 'a': 'Atraso',
    'presente': 'Presente', 'ausente': 'Ausente', 'media clase': 'Media clase', 'atraso': 'Atraso' };
  for (const h of libro.getSheets()) {
    if (!/^asist/i.test(h.getName()) || h.getLastRow() < 2) continue;
    const v = h.getRange(1, 1, h.getLastRow(), h.getLastColumn()).getDisplayValues();
    const fCab = v.findIndex(f => String(f[0]).trim().toUpperCase() === 'RUT');
    if (fCab < 1) continue; // falta la fila 1 con el nombre del curso
    const curso = String(v[0][0]).split('·')[0].trim();
    const fechas = v[fCab].map(fechaTexto);
    for (const f of v.slice(fCab + 1)) {
      if (!f[0]) continue;
      const k = normRut(f[0]);
      const regs = [];
      for (let c = 2; c < f.length; c++) {
        const est = estados[String(f[c]).trim().toLowerCase()];
        if (est && fechas[c]) regs.push({ fecha: fechas[c], estado: est });
      }
      if (regs.length) (datos.asistencia[k] = datos.asistencia[k] || {})[curso] = regs;
    }
  }

  const tipos = { pdf: 'pdf', word: 'doc', powerpoint: 'ppt', video: 'video', audio: 'audio', enlace: 'link' };
  const material = f => ({ titulo: f[1], desc: f[2], tipo: tipos[String(f[3]).toLowerCase()] || 'pdf', url: f[4] });
  datos.materiales = porCurso(tabla('Materiales'), material);
  datos.trabajoGrupo = porCurso(tabla('Trabajo en grupo'), material);
  datos.tareas = porCurso(tabla('Tareas'), f => {
    const t = { titulo: f[1], desc: f[2], fecha: fechaTexto(f[3]) };
    if (f[4]) t.url = f[4];
    return t;
  });
  datos.horarios = {};
  for (const f of tabla('Horarios')) if (f[0]) datos.horarios[f[0]] = { dia: f[1], hora: f[2], modalidad: f[3], lugar: f[4] };
  datos.observacionesCurso = porCurso(tabla('Obs. por curso'), f => ({ fecha: fechaTexto(f[1]), texto: f[2] }));
  for (const f of tabla('Obs. por alumno')) {
    if (!f[0]) continue;
    (datos.observaciones[normRut(f[0])] = datos.observaciones[normRut(f[0])] || []).push({ fecha: fechaTexto(f[2]), texto: f[3] });
  }
  datos.avisos = tabla('Avisos').filter(f => f[3] || f[2]).map(f => ({ fecha: fechaTexto(f[0]), curso: f[1], titulo: f[2], texto: f[3] }));

  try { cache.put('datos', JSON.stringify(datos), 30); } catch (e) { /* demasiado grande para la caché */ }
  return datos;
}

/* "100", "85,5" o vacío → número o null. */
function nota(v) {
  const t = String(v).replace(',', '.').trim();
  if (t === '') return null;
  const n = parseFloat(t);
  return isNaN(n) ? null : n;
}

/* Acepta 6/10/2026, 06-10-2026 o 2026-10-06 y devuelve DD-MM-AAAA. */
function fechaTexto(v) {
  const t = String(v || '').trim();
  let m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) return pad2(m[1]) + '-' + pad2(m[2]) + '-' + (m[3].length === 2 ? '20' + m[3] : m[3]);
  m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return pad2(m[3]) + '-' + pad2(m[2]) + '-' + m[1];
  return t;
}
function pad2(n) { return ('0' + n).slice(-2); }

/* Comprueba que el RUT exista, no esté bloqueado y (si se indica) tenga el curso.
   Usa la planilla de datos si está configurada (y entonces valida también la clave);
   si no, revisa el portal publicado. */
function buscarAlumno(rut, curso, claveAlumno) {
  const datos = leerDatos();
  if (datos) {
    const a = datos.alumnos[normRut(rut)];
    if (!a || a.bloqueado) return null;
    if (a.clave && String(claveAlumno || '').trim() !== a.clave) return null;
    if (curso && !a.cursos[curso]) return null;
    return a.nombre;
  }
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
