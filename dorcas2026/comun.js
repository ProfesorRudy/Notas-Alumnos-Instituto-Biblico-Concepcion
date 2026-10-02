/* =====================================================================
   CONFIGURACIÓN DEL EVENTO
   Cambia estos datos para cada evento.
   ===================================================================== */
const EVENTO = {
  nombre: 'Convención de Dorcas 2026',
  organiza: 'IEPI LAN-C',
  fechas: '',   // por ejemplo: '14 al 16 de noviembre de 2026'
  lugar: '',    // por ejemplo: 'Los Ángeles'
  contacto: '' // teléfono o correo para dudas; se muestra en el formulario y en las tarjetas
};

/* URL de la aplicación web de Google (ver dorcas2026/apps-script.gs).
   Mientras esté vacía, las páginas funcionan en MODO DE PRUEBA: los datos se
   guardan solo en este navegador y la clave del panel es "prueba". */
const EVENTO_URL = '';

const MODO_PRUEBA = !EVENTO_URL;

async function api(accion, datos) {
  const cuerpo = Object.assign({ accion: accion }, datos || {});
  if (MODO_PRUEBA) return apiLocal(cuerpo);
  const r = await fetch(EVENTO_URL, { method: 'POST', body: JSON.stringify(cuerpo) });
  return r.json();
}

/* ---------- Modo de prueba: imita al script de Google con localStorage ---------- */
function apiLocal(d) {
  const CLAVE = 'dorcas2026-prueba';
  let db;
  try { db = JSON.parse(localStorage.getItem(CLAVE)) || null; } catch (e) { db = null; }
  db = db || { inscritos: [], hospedadores: [] };
  const guardar = () => { try { localStorage.setItem(CLAVE, JSON.stringify(db)); } catch (e) {} };
  const siguiente = (lista, prefijo, dig) =>
    prefijo + String(lista.reduce((m, x) => Math.max(m, parseInt(x.codigo.slice(prefijo.length), 10) || 0), 0) + 1).padStart(dig, '0');
  const ok = (extra) => Object.assign({ ok: true }, extra || {});

  if (d.accion === 'inscribir') {
    const codigo = siguiente(db.inscritos, 'I-', 4);
    db.inscritos.push({ codigo, fecha: new Date().toISOString(), nombre: d.nombre, apellido: d.apellido,
      edad: Number(d.edad), iglesia: d.iglesia, telefono: d.telefono || '', hospedador: '',
      movilidad: d.movilidad || '', camaIndividual: d.camaIndividual || '', traslado: d.traslado || '' });
    guardar();
    return ok({ codigo });
  }
  if (d.clave !== 'prueba') return { ok: false, mensaje: 'Clave incorrecta. En modo de prueba la clave es "prueba".' };
  if (d.accion === 'guardarHospedador') {
    const h = d.hospedador;
    const existente = db.hospedadores.find(x => x.codigo === h.codigo);
    const datos = Object.assign({}, h, { capacidad: Number(h.capacidad) || 0 });
    delete datos.codigo;
    if (existente) Object.assign(existente, datos);
    else db.hospedadores.push(Object.assign({ codigo: siguiente(db.hospedadores, 'H-', 2) }, datos));
  } else if (d.accion === 'borrarHospedador') {
    db.hospedadores = db.hospedadores.filter(x => x.codigo !== d.codigo);
    db.inscritos.forEach(i => { if (i.hospedador === d.codigo) i.hospedador = ''; });
  } else if (d.accion === 'borrarInscrito') {
    db.inscritos = db.inscritos.filter(x => x.codigo !== d.codigo);
  } else if (d.accion === 'asignar') {
    db.inscritos.forEach(i => { if (d.codigos.includes(i.codigo)) i.hospedador = d.hospedador || ''; });
  }
  guardar();
  return ok({ datos: db });
}

function esc(t) {
  return String(t == null ? '' : t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
