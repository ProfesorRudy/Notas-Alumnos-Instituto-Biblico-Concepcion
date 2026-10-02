/**
 * Convención de Dorcas 2026 · IEPI LAN-C — inscripciones y hospedaje
 *
 * Guarda en una planilla de Google las inscripciones que llegan desde
 * dorcas2026/index.html y atiende el panel de administración (dorcas2026/admin.html):
 * hospedadores, asignación de hospedaje y datos para las tarjetas.
 *
 * Instalación (una sola vez):
 *   1. Entra a https://script.google.com con tu cuenta de Google y crea un "Proyecto nuevo".
 *   2. Borra el contenido de Código.gs y pega TODO este archivo. Guarda (Ctrl+S).
 *   3. Configuración del proyecto (engranaje) > Propiedades de la secuencia de comandos:
 *      agrega la propiedad CLAVE_ADMIN con la clave que pedirá el panel.
 *   4. Ejecuta la función "probar" una vez (crea la planilla "Convención Dorcas 2026 - Inscripciones"
 *      en tu Drive) y autoriza los permisos que pide.
 *   5. Pulsa "Implementar" > "Nueva implementación" > tipo "Aplicación web".
 *        - Ejecutar como: Yo
 *        - Quién tiene acceso: Cualquier usuario
 *   6. Copia la "URL de la aplicación web" (termina en /exec) y pégala en
 *      EVENTO_URL dentro de dorcas2026/comun.js.
 *
 * La planilla tiene dos hojas que también puedes editar a mano:
 *   - Inscritos:     Código | Fecha | Nombre | Apellido | Edad | Iglesia | Teléfono | Hospedador
 *   - Hospedadores:  Código | Nombre | Dirección | Teléfono | Capacidad | Notas | Camas | Habitaciones |
 *                    Tipo de camas | y seis columnas Sí/No (adultas mayores, primer piso, baño en el mismo
 *                    piso, dificultad para caminar, cama individual, traslado)
 * La columna "Hospedador" de Inscritos guarda el código del hospedador (H-01, H-02...).
 *
 * Para actualizar el código: pega la versión nueva, guarda y luego
 * Implementar > Gestionar implementaciones > Editar (lápiz) > Versión: Nueva versión > Implementar.
 * Así la URL no cambia.
 */

const NOMBRE_PLANILLA = 'Convención Dorcas 2026 - Inscripciones';
const COLS_INSCRITOS = ['Código', 'Fecha', 'Nombre', 'Apellido', 'Edad', 'Iglesia', 'Teléfono', 'Hospedador'];
const COLS_HOSPEDADORES = ['Código', 'Nombre', 'Dirección', 'Teléfono', 'Capacidad', 'Notas', 'Camas', 'Habitaciones',
  'Tipo de camas', 'Adultas mayores', 'Primer piso', 'Baño en el mismo piso', 'Dificultad para caminar',
  'Cama individual', 'Traslado'];
// Preguntas Sí/No del hospedador, en el mismo orden que sus columnas.
const SI_NO = ['adultasMayores', 'primerPiso', 'banoMismoPiso', 'movilidad', 'camaIndividual', 'traslado'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.accion === 'inscribir') return inscribir(d);
    if (!claveValida(d.clave)) return responder(false, 'Clave incorrecta.');
    if (d.accion === 'panel') return responder(true, '', { datos: leerTodo() });
    if (d.accion === 'guardarHospedador') return guardarHospedador(d);
    if (d.accion === 'borrarHospedador') return borrar('Hospedadores', d.codigo);
    if (d.accion === 'borrarInscrito') return borrar('Inscritos', d.codigo);
    if (d.accion === 'asignar') return asignar(d);
    return responder(false, 'Acción desconocida.');
  } catch (err) {
    return responder(false, 'Error: ' + err.message);
  }
}

function doGet() {
  return responder(true, 'Servicio de inscripciones activo.');
}

/* ---------- Inscripción pública ---------- */

function inscribir(d) {
  const nombre = texto(d.nombre, 60);
  const apellido = texto(d.apellido, 60);
  const iglesia = texto(d.iglesia, 100);
  const telefono = texto(d.telefono, 30);
  const edad = parseInt(d.edad, 10);
  if (!nombre || !apellido || !iglesia) return responder(false, 'Faltan datos obligatorios.');
  if (!(edad >= 0 && edad <= 120)) return responder(false, 'La edad no es válida.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const hoja = hojaDe('Inscritos');
    const codigo = siguienteCodigo(hoja, 'I-', 4);
    hoja.appendRow([codigo, new Date(), nombre, apellido, edad, iglesia, telefono, '']);
    return responder(true, 'Inscripción recibida.', { codigo: codigo });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Panel de administración ---------- */

function leerTodo() {
  const inscritos = filas('Inscritos').map(f => ({
    codigo: f[0], fecha: f[1] instanceof Date ? f[1].toISOString() : String(f[1]),
    nombre: f[2], apellido: f[3], edad: f[4], iglesia: f[5], telefono: String(f[6]), hospedador: f[7]
  }));
  const hospedadores = filas('Hospedadores').map(f => {
    const h = { codigo: f[0], nombre: f[1], direccion: f[2], telefono: String(f[3]), capacidad: Number(f[4]) || 0, notas: f[5],
                camas: f[6] === '' || f[6] == null ? '' : Number(f[6]), habitaciones: f[7] === '' || f[7] == null ? '' : Number(f[7]),
                tipoCamas: f[8] || '' };
    SI_NO.forEach((k, i) => h[k] = siNo(f[9 + i]));
    return h;
  });
  return { inscritos: inscritos, hospedadores: hospedadores };
}

function guardarHospedador(d) {
  const h = d.hospedador || {};
  const fila = [texto(h.nombre, 80), texto(h.direccion, 150), texto(h.telefono, 30),
                Math.max(0, parseInt(h.capacidad, 10) || 0), texto(h.notas, 200),
                numeroOpcional(h.camas), numeroOpcional(h.habitaciones),
                ['Individuales', 'Matrimoniales', 'Ambas'].indexOf(h.tipoCamas) >= 0 ? h.tipoCamas : '']
                .concat(SI_NO.map(k => siNo(h[k])));
  if (!fila[0]) return responder(false, 'El hospedador necesita un nombre.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const hoja = hojaDe('Hospedadores');
    let codigo = String(h.codigo || '');
    const n = codigo ? buscarFila(hoja, codigo) : 0;
    if (n) {
      hoja.getRange(n, 2, 1, fila.length).setValues([fila]);
    } else {
      codigo = siguienteCodigo(hoja, 'H-', 2);
      hoja.appendRow([codigo].concat(fila));
    }
    return responder(true, 'Hospedador guardado.', { datos: leerTodo() });
  } finally {
    lock.releaseLock();
  }
}

function asignar(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const hoja = hojaDe('Inscritos');
    const destino = String(d.hospedador || '');
    (d.codigos || []).forEach(codigo => {
      const n = buscarFila(hoja, codigo);
      if (n) hoja.getRange(n, 8).setValue(destino);
    });
    return responder(true, 'Asignación guardada.', { datos: leerTodo() });
  } finally {
    lock.releaseLock();
  }
}

function borrar(nombreHoja, codigo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const hoja = hojaDe(nombreHoja);
    const n = buscarFila(hoja, codigo);
    if (n) hoja.deleteRow(n);
    // Si se borra un hospedador, sus huéspedes quedan sin asignar.
    if (nombreHoja === 'Hospedadores') {
      const ins = hojaDe('Inscritos');
      const ultima = ins.getLastRow();
      if (ultima > 1) {
        const r = ins.getRange(2, 8, ultima - 1, 1);
        r.setValues(r.getValues().map(v => [v[0] === codigo ? '' : v[0]]));
      }
    }
    return responder(true, 'Eliminado.', { datos: leerTodo() });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Planilla ---------- */

function planilla() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PLANILLA_EVENTO');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create(NOMBRE_PLANILLA);
  const ins = ss.getSheets()[0];
  ins.setName('Inscritos');
  ins.appendRow(COLS_INSCRITOS);
  ins.setFrozenRows(1);
  const hos = ss.insertSheet('Hospedadores');
  hos.appendRow(COLS_HOSPEDADORES);
  hos.setFrozenRows(1);
  props.setProperty('PLANILLA_EVENTO', ss.getId());
  return ss;
}

function hojaDe(nombre) {
  const hoja = planilla().getSheetByName(nombre);
  // Una planilla creada con una versión anterior recibe aquí las columnas nuevas.
  if (nombre === 'Hospedadores' && hoja.getLastColumn() < COLS_HOSPEDADORES.length) {
    hoja.getRange(1, 1, 1, COLS_HOSPEDADORES.length).setValues([COLS_HOSPEDADORES]);
  }
  return hoja;
}

function filas(nombre) {
  const hoja = hojaDe(nombre);
  const ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  return hoja.getRange(2, 1, ultima - 1, hoja.getLastColumn()).getValues().filter(f => f[0] !== '');
}

function buscarFila(hoja, codigo) {
  const ultima = hoja.getLastRow();
  if (ultima < 2 || !codigo) return 0;
  const codigos = hoja.getRange(2, 1, ultima - 1, 1).getValues();
  for (let i = 0; i < codigos.length; i++) if (codigos[i][0] === codigo) return i + 2;
  return 0;
}

function siguienteCodigo(hoja, prefijo, digitos) {
  let max = 0;
  const ultima = hoja.getLastRow();
  if (ultima > 1) {
    hoja.getRange(2, 1, ultima - 1, 1).getValues().forEach(v => {
      const n = parseInt(String(v[0]).replace(prefijo, ''), 10);
      if (n > max) max = n;
    });
  }
  return prefijo + String(max + 1).padStart(digitos, '0');
}

/* ---------- Utilidades ---------- */

function claveValida(clave) {
  const real = PropertiesService.getScriptProperties().getProperty('CLAVE_ADMIN');
  return !!real && String(clave || '') === real;
}

// Recorta y evita que un texto se interprete como fórmula en la planilla
// (el apóstrofo inicial obliga a guardarlo como texto, p. ej. "+56 9 ...").
function texto(v, max) {
  const t = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function siNo(v) { return v === 'Sí' || v === 'No' ? v : ''; }

function numeroOpcional(v) {
  const n = parseInt(v, 10);
  return n >= 0 ? n : '';
}

function responder(ok, mensaje, extra) {
  return ContentService.createTextOutput(JSON.stringify(Object.assign({ ok: ok, mensaje: mensaje }, extra || {})))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Prueba rápida desde el editor: Ejecutar > probar. Crea la planilla y muestra su dirección. */
function probar() {
  Logger.log('Planilla: ' + planilla().getUrl());
}
