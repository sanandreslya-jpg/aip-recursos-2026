/**
 * SIGAIP - Reservas de recursos tecnológicos (Aula de Innovación Pedagógica)
 * I.E.N.° 16210 "Alejandro Sánchez Arteaga"
 *
 * Backend en Google Apps Script + Google Sheets.
 * Ejecutar UNA VEZ la función setup() para crear las hojas con datos de ejemplo.
 */

const DIAS = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
const H = { RECURSOS: 'Recursos', DOCENTES: 'Docentes', BLOQUES: 'Bloques', RESERVAS: 'Reservas' };
const COLS_RESERVA = ['ID', 'Registro', 'Fecha', 'Día', 'Bloque', 'Inicio', 'Fin', 'Recurso', 'Cantidad',
  'Docente', 'Área', 'Grado y Sección', 'Requerimientos', 'Estado'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SIGAIP - Reservas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- Configuración inicial ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const crear = (nombre, cabecera, filas, textoCols) => {
    let sh = ss.getSheetByName(nombre);
    if (sh) return;
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cabecera.length).setValues([cabecera]).setFontWeight('bold').setBackground('#dbeafe');
    sh.setFrozenRows(1);
    (textoCols || []).forEach(c => sh.getRange(2, c, 1000, 1).setNumberFormat('@')); // evita que Sheets convierta a fecha/hora
    if (filas.length) sh.getRange(2, 1, filas.length, cabecera.length).setValues(filas);
  };
  crear(H.RECURSOS, ['Recurso', 'Cantidad', 'Icono', 'Descripción'], [
    ['Laptops', 35, '💻', 'Se reservan por cantidad'],
    ['Proyector', 7, '📽️', ''],
    ['Parlante', 4, '🔊', ''],
    ['Sala de cómputo', 1, '🖥️', 'Aula completa']
  ]);
  crear(H.DOCENTES, ['Docente', 'Área curricular', 'Activo'], [
    ['Docente de ejemplo 1', 'Matemática', 'SI'],
    ['Docente de ejemplo 2', 'Comunicación', 'SI']
  ]);
  // Edita estos bloques según el horario real de tu I.E.
  crear(H.BLOQUES, ['Bloque', 'Inicio', 'Fin'], [
    ['1.ª hora', '07:45', '08:30'], ['2.ª hora', '08:30', '09:15'], ['3.ª hora', '09:15', '10:00'],
    ['Recreo', '10:00', '10:30'],
    ['4.ª hora', '10:30', '11:15'], ['5.ª hora', '11:15', '12:00'], ['6.ª hora', '12:00', '12:45']
  ], [2, 3]);
  crear(H.RESERVAS, COLS_RESERVA, [], [1, 3, 6, 7]);
  const s1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
}

/* ---------- Lectura de datos ---------- */

function tabla_(nombre) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  const v = sh.getDataRange().getDisplayValues();
  const cab = v.shift();
  return v.filter(f => f.some(c => c !== '')).map(f => {
    const o = {}; cab.forEach((c, i) => o[c] = f[i]); return o;
  });
}

function hoy_() { return Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd'); }
function ahora_() { return Utilities.formatDate(new Date(), 'America/Lima', 'HH:mm'); }

/** Días hábiles (lun-vie) desde hoy hasta el viernes de la próxima semana. */
function fechasDisponibles_() {
  const out = [];
  const d = new Date();
  d.setHours(12, 0, 0, 0);
  const limite = ((5 - d.getDay() + 7) % 7) + 7; // hasta el viernes de la próxima semana
  for (let i = 0; i <= limite; i++) {
    const x = new Date(d.getTime() + i * 86400000);
    if (x.getDay() >= 1 && x.getDay() <= 5)
      out.push({ fecha: Utilities.formatDate(x, 'America/Lima', 'yyyy-MM-dd'), dia: DIAS[x.getDay()] });
  }
  return out;
}

function getInit() {
  return {
    hoy: hoy_(),
    recursos: tabla_(H.RECURSOS).map(r => ({ nombre: r['Recurso'], total: +r['Cantidad'], icono: r['Icono'], desc: r['Descripción'] })),
    docentes: tabla_(H.DOCENTES).filter(d => (d['Activo'] || 'SI').toUpperCase() !== 'NO')
      .map(d => ({ nombre: d['Docente'], area: d['Área curricular'] }))
      .sort((a, b) => a.nombre.localeCompare(b.nombre)),
    fechas: fechasDisponibles_()
  };
}

/** Para una fecha y un recurso: cuántas unidades libres hay en cada bloque. */
function getDisponibilidad(fecha, recurso) {
  const rec = tabla_(H.RECURSOS).find(r => r['Recurso'] === recurso);
  const total = +rec['Cantidad'];
  const usadas = {};
  tabla_(H.RESERVAS).filter(r => r['Fecha'] === fecha && r['Recurso'] === recurso && r['Estado'] === 'Activa')
    .forEach(r => usadas[r['Bloque']] = (usadas[r['Bloque']] || 0) + (+r['Cantidad']));
  const esHoy = fecha === hoy_(), ahora = ahora_();
  return tabla_(H.BLOQUES).map(b => ({
    bloque: b['Bloque'], inicio: b['Inicio'], fin: b['Fin'],
    esRecreo: /recreo|receso|almuerzo/i.test(b['Bloque']),
    pasado: esHoy && b['Fin'] <= ahora,
    libres: total - (usadas[b['Bloque']] || 0),
    total: total
  }));
}

/* ---------- Reservas ---------- */

function crearReserva(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita que dos docentes reserven lo mismo a la vez
  try {
    if (!p.docente || !p.fecha || !p.recurso || !p.bloques || !p.bloques.length || !p.grado)
      throw new Error('Completa todos los campos obligatorios.');
    if (!fechasDisponibles_().some(f => f.fecha === p.fecha))
      throw new Error('Esa fecha no está habilitada para reservar.');
    const docente = tabla_(H.DOCENTES).find(d => d['Docente'] === p.docente);
    if (!docente) throw new Error('Docente no válido.');
    const cant = Math.max(1, parseInt(p.cantidad, 10) || 1);
    const disp = getDisponibilidad(p.fecha, p.recurso);
    p.bloques.forEach(nombre => {
      const b = disp.find(x => x.bloque === nombre);
      if (!b) throw new Error('Bloque no válido: ' + nombre);
      if (b.pasado) throw new Error(nombre + ' ya pasó.');
      if (b.libres < cant) throw new Error('En "' + nombre + '" solo quedan ' + Math.max(b.libres, 0) + ' disponibles.');
    });
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.RESERVAS);
    const fechaDia = DIAS[new Date(p.fecha + 'T12:00:00').getDay()];
    const id = Utilities.getUuid().slice(0, 8);
    const filas = p.bloques.map(nombre => {
      const b = disp.find(x => x.bloque === nombre);
      return [id, Utilities.formatDate(new Date(), 'America/Lima', 'yyyy-MM-dd HH:mm'), p.fecha, fechaDia, nombre,
        b.inicio, b.fin, p.recurso, cant, p.docente, docente['Área curricular'], p.grado, p.notas || '', 'Activa'];
    });
    sh.getRange(sh.getLastRow() + 1, 1, filas.length, COLS_RESERVA.length).setValues(filas);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function misReservas(docente) {
  const hoy = hoy_();
  return tabla_(H.RESERVAS)
    .filter(r => r['Docente'] === docente && r['Estado'] === 'Activa' && r['Fecha'] >= hoy)
    .sort((a, b) => (a['Fecha'] + a['Inicio']).localeCompare(b['Fecha'] + b['Inicio']))
    .map(r => ({ id: r['ID'], fecha: r['Fecha'], dia: r['Día'], bloque: r['Bloque'], inicio: r['Inicio'], fin: r['Fin'],
      recurso: r['Recurso'], cantidad: r['Cantidad'], grado: r['Grado y Sección'] }));
}

function cancelarReserva(id, docente) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H.RESERVAS);
    const v = sh.getDataRange().getDisplayValues();
    const iId = v[0].indexOf('ID'), iDoc = v[0].indexOf('Docente'), iEst = v[0].indexOf('Estado');
    let n = 0;
    for (let i = 1; i < v.length; i++) {
      if (v[i][iId] === id && v[i][iDoc] === docente && v[i][iEst] === 'Activa') {
        sh.getRange(i + 1, iEst + 1).setValue('Cancelada'); n++;
      }
    }
    return { ok: n > 0 };
  } finally {
    lock.releaseLock();
  }
}
