/**
 * Backend del AIP 16210: guarda inventario y movimientos en la Hoja de Google
 * a la que está vinculado este script. Publicar como Aplicación web:
 *   Ejecutar como: Yo   |   Quién tiene acceso: Cualquier persona
 */
const INV_H = ['cod', 'tipo', 'desc', 'estado'];
const MOV_H = ['id', 'nombre', 'dni', 'rol', 'grado', 'area', 'obs', 'cods', 'fecha', 'limite', 'devuelto', 'estadoDev', 'obsDev'];

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.getRange(1, 1, s.getMaxRows(), headers.length).setNumberFormat('@'); // todo como texto (DNI, fechas)
    s.appendRow(headers);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return s;
}

function rows_(s, headers) {
  const v = s.getDataRange().getValues();
  return v.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    headers.forEach((k, i) => (o[k] = String(r[i])));
    return o;
  });
}

function state_() {
  const inv = rows_(sheet_('Inventario', INV_H), INV_H);
  const mov = rows_(sheet_('Movimientos', MOV_H), MOV_H).map(m => ({
    id: m.id, nombre: m.nombre, dni: m.dni, rol: m.rol, grado: m.grado, area: m.area, obs: m.obs,
    cods: m.cods ? m.cods.split(' ') : [],
    fecha: m.fecha, limite: m.limite,
    devuelto: m.devuelto || null, estadoDev: m.estadoDev || undefined, obsDev: m.obsDev || ''
  }));
  return { inv: inv, mov: mov };
}

function findRow_(s, key) {
  const col = s.getRange(1, 1, s.getLastRow(), 1).getValues();
  for (let i = 1; i < col.length; i++) if (String(col[i][0]) === String(key)) return i + 1;
  return 0;
}

function setEstado_(cod, estado) {
  const s = sheet_('Inventario', INV_H), r = findRow_(s, cod);
  if (r) s.getRange(r, 4).setValue(estado);
}

function apply_(e) {
  const inv = sheet_('Inventario', INV_H), mov = sheet_('Movimientos', MOV_H);
  switch (e.action) {
    case 'init':
      if (inv.getLastRow() <= 1) e.inv.forEach(x => inv.appendRow([x.cod, x.tipo, x.desc, x.estado]));
      break;
    case 'addItem':
      if (!findRow_(inv, e.item.cod)) inv.appendRow([e.item.cod, e.item.tipo, e.item.desc, 'disp']);
      break;
    case 'delItem': {
      const r = findRow_(inv, e.cod);
      if (r) inv.deleteRow(r);
      break;
    }
    case 'setEstado':
      setEstado_(e.cod, e.estado);
      break;
    case 'addMov': {
      const m = e.m;
      mov.appendRow([m.id, m.nombre, m.dni, m.rol, m.grado, m.area, m.obs, m.cods.join(' '), m.fecha, m.limite, '', '', '']);
      m.cods.forEach(c => setEstado_(c, 'pres'));
      break;
    }
    case 'ret': {
      const r = findRow_(mov, e.id);
      if (r) {
        const cods = String(mov.getRange(r, 8).getValue()).split(' ');
        mov.getRange(r, 11, 1, 3).setValues([[e.devuelto, e.estadoDev, e.obsDev || '']]);
        cods.forEach(c => setEstado_(c, e.estadoDev === 'dano' ? 'mant' : 'disp'));
      }
      break;
    }
    case 'reset':
      [[inv, 4], [mov, MOV_H.length]].forEach(p => { if (p[0].getLastRow() > 1) p[0].deleteRows(2, p[0].getLastRow() - 1); });
      e.inv.forEach(x => inv.appendRow([x.cod, x.tipo, x.desc, x.estado]));
      break;
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  try { return json_(state_()); } catch (err) { return json_({ error: String(err) }); }
}

function doPost(ev) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    apply_(JSON.parse(ev.postData.contents));
    SpreadsheetApp.flush();
    return json_(state_());
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
