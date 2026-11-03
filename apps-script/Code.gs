/**
 * Caja Criolla — conector con Google Sheets.
 *
 * 1. Cambia CLAVE por una palabra secreta tuya (la misma que pondrás en la app).
 * 2. Ejecuta la función "preparar" una vez (crea las pestañas).
 * 3. Implementar → Nueva implementación → Aplicación web
 *    · Ejecutar como: Yo   · Quién tiene acceso: Cualquier usuario
 * 4. Copia la URL que termina en /exec y pégala en la app (Más → Conexión).
 */
const CLAVE = 'CAMBIA-ESTA-CLAVE';
const ZONA = 'America/Puerto_Rico';

const HOJAS = {
  ventas: { nombre: 'Ventas', cols: ['Fecha', 'Producto', 'Cantidad', 'Método de pago', 'Precio unitario', 'Total', 'Tipo', 'Arroz', 'Granos', 'Carne', 'Side', 'IVU %', 'IVU $', 'Hora', 'Equipo', 'ID', 'Marca'] },
  gastos: { nombre: 'Gastos', cols: ['Fecha', 'Categoría', 'Descripción', 'Suplidor', 'Monto', 'Método de pago', 'Hora', 'Equipo', 'ID', 'Marca'] },
  merma:  { nombre: 'Merma',  cols: ['Fecha', 'Componente', 'Porciones perdidas', 'Razón', 'Hora', 'Equipo', 'ID', 'Marca'] },
  menu:   { nombre: 'Menú del día', cols: ['Fecha', 'Menú', 'Actualizado', 'Equipo'] },
  borrados: { nombre: 'Borrados', cols: ['Borrado el', 'Hoja', 'Fecha', 'Detalle', 'Equipo', 'ID'] },
  config: { nombre: 'Config', cols: ['Clave', 'Valor (no editar a mano)', 'Actualizado'] }
};

function preparar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HOJAS).forEach(k => hoja_(ss, k));
  const vacia = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (vacia && vacia.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(vacia);
  return 'Listo';
}

function hoja_(ss, k) {
  const h = HOJAS[k];
  let sh = ss.getSheetByName(h.nombre);
  if (!sh) {
    sh = ss.insertSheet(h.nombre);
    sh.getRange(1, 1, 1, h.cols.length).setValues([h.cols]).setFontWeight('bold').setBackground('#E8F1FA');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    const iId = h.cols.indexOf('ID');
    if (iId >= 0) sh.getRange(1, iId + 1, sh.getMaxRows(), 1).setNumberFormat('@');
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.clave !== CLAVE) return json_({ ok: false, error: 'clave' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (p.accion === 'ping') return json_({ ok: true, hoja: ss.getName() });
    if (p.accion === 'config') return json_({ ok: true, config: leerConfig_(ss) });
    if (p.accion === 'dias') return json_({ ok: true, dias: leerDias_(ss, p.desde, p.hasta), config: p.conConfig ? leerConfig_(ss) : undefined });
    return json_({ ok: false, error: 'accion' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.clave !== CLAVE) return json_({ ok: false, error: 'clave' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipo = String(body.equipo || '').slice(0, 40);
    const hechos = [];
    const idsCache = {};
    (body.ops || []).forEach(op => {
      if (op.op === 'add') agregar_(ss, op, equipo, idsCache);
      else if (op.op === 'del') borrar_(ss, op, equipo, idsCache);
      else if (op.op === 'menu') menuDia_(ss, op, equipo);
      else if (op.op === 'config') guardarConfig_(ss, op.data);
      hechos.push(op.oid);
    });
    SpreadsheetApp.flush();
    return json_({ ok: true, hechos: hechos });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function ids_(sh, k, cache) {
  if (cache[k]) return cache[k];
  const col = HOJAS[k].cols.indexOf('ID') + 1;
  const n = sh.getLastRow() - 1;
  const vals = n > 0 ? sh.getRange(2, col, n, 1).getValues().map(r => String(r[0])) : [];
  cache[k] = vals;
  return vals;
}

function hora_(t) {
  return t ? Utilities.formatDate(new Date(Number(t)), ZONA, 'h:mm a') : '';
}

function agregar_(ss, op, equipo, cache) {
  const k = op.hoja;
  if (!HOJAS[k] || ['ventas', 'gastos', 'merma'].indexOf(k) < 0) return;
  const sh = hoja_(ss, k);
  const ids = ids_(sh, k, cache);
  if (ids.indexOf(String(op.id)) >= 0) return; // ya estaba (reintento)
  const d = op.data || {};
  let fila;
  if (k === 'ventas') {
    const cant = Number(d.cant) || 0, precio = Number(d.precio) || 0, tasa = Number(d.ivu) || 0;
    fila = [op.fecha, d.plato || '', cant, d.metodo || '', precio, round2_(cant * precio), d.tipo || '', d.arroz || '', d.granos || '', d.carne || '', d.side || '', tasa, round2_(cant * precio * tasa), hora_(d.t), equipo, String(op.id), Number(d.t) || ''];
  } else if (k === 'gastos') {
    fila = [op.fecha, d.cat || '', d.desc || '', d.sup || '', Number(d.monto) || 0, d.metodo || '', hora_(d.t), equipo, String(op.id), Number(d.t) || ''];
  } else {
    fila = [op.fecha, d.comp || '', Number(d.porc) || 0, d.razon || '', hora_(d.t), equipo, String(op.id), Number(d.t) || ''];
  }
  sh.appendRow(fila);
  ids.push(String(op.id));
}

function borrar_(ss, op, equipo, cache) {
  const k = op.hoja;
  if (['ventas', 'gastos', 'merma'].indexOf(k) < 0) return;
  const sh = hoja_(ss, k);
  const ids = ids_(sh, k, cache);
  const i = ids.indexOf(String(op.id));
  if (i < 0) return;
  const row = i + 2;
  const vals = sh.getRange(row, 1, 1, HOJAS[k].cols.length).getValues()[0];
  hoja_(ss, 'borrados').appendRow([Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd h:mm a'), HOJAS[k].nombre, vals[0], vals.slice(1, 6).join(' · '), equipo, String(op.id)]);
  sh.deleteRow(row);
  ids.splice(i, 1);
}

function menuDia_(ss, op, equipo) {
  const sh = hoja_(ss, 'menu');
  const n = sh.getLastRow() - 1;
  const fechas = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0])) : [];
  const fila = [op.fecha, (op.menu || []).join(', '), Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd h:mm a'), equipo];
  const i = fechas.indexOf(String(op.fecha));
  if (i >= 0) sh.getRange(i + 2, 1, 1, fila.length).setValues([fila]);
  else sh.appendRow(fila);
}

function guardarConfig_(ss, data) {
  const sh = hoja_(ss, 'config');
  const fila = ['menu', JSON.stringify(data || {}), Utilities.formatDate(new Date(), ZONA, 'yyyy-MM-dd h:mm a')];
  sh.getRange(2, 1, 1, 3).setValues([fila]);
}

function leerConfig_(ss) {
  const sh = ss.getSheetByName(HOJAS.config.nombre);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 2).getValue();
  try { return v ? JSON.parse(v) : null; } catch (e) { return null; }
}

function texto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, ZONA, 'yyyy-MM-dd');
  return String(v);
}

function leerDias_(ss, desde, hasta) {
  const out = {};
  const dia = f => out[f] || (out[f] = { fecha: f, ventas: {}, gastos: {}, merma: {} });
  const enRango = f => f >= desde && f <= hasta;
  ['ventas', 'gastos', 'merma'].forEach(k => {
    const sh = ss.getSheetByName(HOJAS[k].nombre);
    if (!sh || sh.getLastRow() < 2) return;
    const cols = HOJAS[k].cols;
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, cols.length).getValues();
    const c = n => cols.indexOf(n);
    vals.forEach(r => {
      const f = texto_(r[0]);
      if (!enRango(f)) return;
      const id = String(r[c('ID')]);
      const t = Number(r[c('Marca')]) || 0;
      let e;
      if (k === 'ventas') e = { t: t, tipo: r[c('Tipo')] || 'plato', plato: r[c('Producto')], arroz: r[c('Arroz')], granos: r[c('Granos')], carne: r[c('Carne')], side: r[c('Side')], cant: Number(r[c('Cantidad')]) || 0, precio: Number(r[c('Precio unitario')]) || 0, metodo: r[c('Método de pago')], ivu: Number(r[c('IVU %')]) || 0 };
      else if (k === 'gastos') e = { t: t, cat: r[c('Categoría')], desc: r[c('Descripción')], sup: r[c('Suplidor')], monto: Number(r[c('Monto')]) || 0, metodo: r[c('Método de pago')] };
      else e = { t: t, comp: r[c('Componente')], porc: Number(r[c('Porciones perdidas')]) || 0, razon: r[c('Razón')] };
      dia(f)[k][id] = e;
    });
  });
  const sm = ss.getSheetByName(HOJAS.menu.nombre);
  if (sm && sm.getLastRow() >= 2) {
    sm.getRange(2, 1, sm.getLastRow() - 1, 2).getValues().forEach(r => {
      const f = texto_(r[0]);
      if (enRango(f)) dia(f).menu = String(r[1]).split(', ').filter(String);
    });
  }
  return out;
}

function round2_(n) { return Math.round(n * 100) / 100; }
