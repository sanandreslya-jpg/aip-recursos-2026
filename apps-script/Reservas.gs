/**
 * Reservas semanales de recursos del AIP 16210.
 * Se agrega al mismo proyecto de Apps Script que Code.gs (mismo enlace /exec).
 * Hojas que crea solas: "Docentes" (nombre, area, clave) y "Reservas".
 */
const ADMIN_PIN = '1234'; // CAMBIE este PIN: lo usa la coordinación del AIP para anular reservas y bloquear horarios

// ==CORE== (este bloque también está copiado en reservas.html; mantenerlos iguales)
const CAP = { laptops: 35, parlantes: 5, proyectores: 5, sala: 1 };
const ITEMS = ['laptops', 'parlantes', 'proyectores', 'sala'];
const NOMBRES = { laptops: 'laptops', parlantes: 'parlantes', proyectores: 'proyectores', sala: 'sala de cómputo' };
const MAX_BLOQUES = 15; // bloques por docente y por semana
const HORAS = 7;

function addDays(s, n) { const d = new Date(s + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function lunesDe(s) { const d = new Date(s + 'T12:00:00Z'); return addDays(s, -((d.getUTCDay() + 6) % 7)); }
function semanaBase(hoy) { return lunesDe(addDays(hoy, 2)); } // sábado y domingo ya miran la semana siguiente
function usoDe(reservas, fecha, hora) {
  const u = { laptops: 0, parlantes: 0, proyectores: 0, sala: 0 };
  reservas.forEach(r => { if (r.fecha === fecha && +r.hora === +hora) ITEMS.forEach(k => (u[k] += +r[k] || 0)); });
  return u;
}
function validar(reservas, r, hoy) {
  if (!r.docente) return 'Elija su nombre';
  const dow = new Date(r.fecha + 'T12:00:00Z').getUTCDay();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.fecha) || dow < 1 || dow > 5) return 'Fecha no válida';
  if (r.fecha < hoy) return 'No se puede reservar en una fecha pasada';
  if (r.fecha > addDays(semanaBase(hoy), 11)) return 'Solo puede reservar esta semana y la siguiente';
  if (!(+r.hora >= 1 && +r.hora <= HORAS)) return 'Hora no válida';
  let total = 0;
  for (const k of ITEMS) {
    const v = +r[k] || 0;
    if (v < 0 || v !== Math.floor(v)) return 'Cantidad no válida';
    total += v;
  }
  if (!total) return 'Elija al menos un recurso';
  const u = usoDe(reservas, r.fecha, r.hora);
  for (const k of ITEMS) {
    if ((+r[k] || 0) > CAP[k] - u[k]) return 'No alcanza: en ese bloque quedan ' + (CAP[k] - u[k]) + ' de ' + NOMBRES[k] + '. Actualice y vuelva a intentar';
  }
  if (!r.bloqueo) {
    if (reservas.some(x => x.docente === r.docente && x.fecha === r.fecha && +x.hora === +r.hora)) return 'Ya tiene una reserva en ese bloque';
    const sem = lunesDe(r.fecha);
    if (reservas.filter(x => x.docente === r.docente && lunesDe(x.fecha) === sem).length >= MAX_BLOQUES) return 'Llegó al máximo de ' + MAX_BLOQUES + ' bloques por semana';
  }
  return null;
}
function validarVarias(reservas, rs, hoy) {
  const acc = reservas.slice();
  for (const r of rs) {
    const e = validar(acc, r, hoy);
    if (e) return e + ' (' + r.hora + '.ª hora)';
    acc.push(r);
  }
  return null;
}
function planBloqueo(reservas, o, hoy) {
  const rows = [];
  for (let h = +o.desde; h <= +o.hasta; h++) {
    const u = usoDe(reservas.concat(rows), o.fecha, h);
    const r = { fecha: o.fecha, hora: h, docente: 'BLOQUEO', area: '', grado: '', nota: o.motivo || 'Bloqueado', bloqueo: true };
    let t = 0;
    ITEMS.forEach(k => { r[k] = CAP[k] - u[k]; t += r[k]; });
    if (!t) continue;
    const e = validar(reservas.concat(rows), r, hoy);
    if (e) return { error: e };
    rows.push(r);
  }
  return { rows: rows };
}
// ==/CORE==

const RES_H = ['id', 'fecha', 'hora', 'docente', 'area', 'laptops', 'parlantes', 'proyectores', 'sala', 'grado', 'nota', 'creado'];
const DOC_H = ['nombre', 'area', 'clave'];
const RES_ACTIONS = ['initDocentes', 'reservar', 'reservarVarias', 'cancelar', 'bloquear', 'cancelarAdmin', 'adminCheck'];

function hoyLima_() { return Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd'); }

function resRead_() {
  return rows_(sheet_('Reservas', RES_H), RES_H).map(r => ({
    id: r.id, fecha: r.fecha, hora: +r.hora, docente: r.docente, area: r.area,
    laptops: +r.laptops || 0, parlantes: +r.parlantes || 0, proyectores: +r.proyectores || 0, sala: +r.sala || 0,
    grado: r.grado, nota: r.nota, creado: r.creado
  }));
}

function resState_(extra) {
  const hoy = hoyLima_(), desde = addDays(lunesDe(hoy), -7);
  const docentes = rows_(sheet_('Docentes', DOC_H), DOC_H).map(d => ({ nombre: d.nombre, area: d.area, protegido: !!d.clave }));
  return Object.assign({ docentes: docentes, reservas: resRead_().filter(r => r.fecha >= desde), hoy: hoy }, extra || {});
}

function reservasGet_() { return resState_(); }

function appendRes_(sh, r) {
  sh.appendRow([Utilities.getUuid().slice(0, 8), r.fecha, r.hora, r.docente, r.area || '', r.laptops || 0, r.parlantes || 0,
    r.proyectores || 0, r.sala || 0, r.grado || '', r.nota || '', new Date().toISOString()]);
}

function reservasPost_(e) {
  const hoy = hoyLima_();
  const docSh = sheet_('Docentes', DOC_H), sh = sheet_('Reservas', RES_H);
  const fail = m => resState_({ error: m });
  const admin = e.pin === ADMIN_PIN;
  switch (e.action) {
    case 'initDocentes':
      if (docSh.getLastRow() <= 1) e.docentes.forEach(d => docSh.appendRow([d.nombre, d.area, '']));
      break;
    case 'adminCheck':
      if (!admin) return fail('PIN incorrecto');
      break;
    case 'reservarVarias': {
      const rs = e.rs || [];
      if (!rs.length || rs.length > 7) return fail('Solicitud no válida');
      const d = rows_(docSh, DOC_H).find(x => x.nombre === rs[0].docente);
      if (!d || rs.some(r => r.docente !== d.nombre)) return fail('Docente no registrado');
      if (d.clave && String(e.clave || '') !== String(d.clave)) return fail('Clave incorrecta');
      rs.forEach(r => (r.bloqueo = false));
      const err = validarVarias(resRead_(), rs, hoy);
      if (err) return fail(err);
      rs.forEach(r => appendRes_(sh, r));
      break;
    }
    case 'cancelar': {
      const d = rows_(docSh, DOC_H).find(x => x.nombre === e.docente);
      if (d && d.clave && String(e.clave || '') !== String(d.clave)) return fail('Clave incorrecta');
      for (const id of (e.ids || [e.id])) {
        const row = findRow_(sh, id);
        if (!row || sh.getRange(row, 4).getValue() !== e.docente) return fail('No se encontró la reserva');
        sh.deleteRow(row);
      }
      break;
    }
    case 'cancelarAdmin': {
      if (!admin) return fail('PIN incorrecto');
      (e.ids || [e.id]).forEach(id => { const row = findRow_(sh, id); if (row) sh.deleteRow(row); });
      break;
    }
    case 'bloquear': {
      if (!admin) return fail('PIN incorrecto');
      const plan = planBloqueo(resRead_(), e, hoy);
      if (plan.error) return fail(plan.error);
      plan.rows.forEach(r => appendRes_(sh, r));
      break;
    }
  }
  SpreadsheetApp.flush();
  return resState_();
}
