/**
 * Backend de Preventivos UDES.
 *
 * Al SUBIR el cronograma (modo: 'cargar'):
 *   - borra TODAS las hojas actuales del Google Sheet;
 *   - crea una sola hoja con el nombre de la hoja elegida en el Excel (2026-B, 2026-A...);
 *   - copia valores, colores de fondo, fuentes, combinaciones, anchos y altos;
 *   - aplica encima los estados ya registrados en la app.
 *
 * Después (modo: 'fila' / 'total') escribe solo las celdas de cada equipo,
 * sin crear hojas auxiliares ni "Informe".
 *
 * 1) Google Sheet → Extensiones → Apps Script → pega este código.
 * 2) Ejecuta setup() una vez y autoriza.
 * 3) Implementar > Nueva implementación > Aplicación web >
 *    Ejecutar como: Yo > Acceso: Cualquier persona con el enlace.
 * 4) Pega la URL en Datos en la app (clave: udes2026).
 */
const TOKEN = 'udes2026';
const FONDOS = {
  OK: '#92D050',
  NOVEDAD: '#FFC000',
  AUSENTE: '#FF0000',
  ALTA: '#00B0F0',
  PENDIENTE: '#E9EEF4',
};
const HEADERS = [
  'EDIFICIO', 'PISO', 'OFICINA', 'SERIAL UDES', 'TIPO EQUIPO', 'FECHA PROGRAMADA',
  'FECHA EJECUTADA', 'TÉCNICO', 'SEMESTRE', 'TIEMPO EJECUCIÓN', 'FECHA SEGUIMIENTO', 'OBSERVACIONES',
];
const PATRONES = {
  edificio: /edificio|sede|bloque/i,
  piso: /piso|nivel/i,
  oficina: /oficina|dependencia|area/i,
  serial: /serial|inventario/i,
  tipo: /tipo.*equipo|equipo/i,
  programada: /fecha.*programada|programad/i,
  ejecutada: /fecha.*ejecutada|ejecutad/i,
  tecnico: /tecnico|técnico/i,
  semestre: /semestre/i,
  duracion: /tiempo|duracion|duración/i,
  seguimiento: /seguimiento/i,
  observaciones: /observacion|observación|novedad/i,
};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Preventivos UDES');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  return 'Listo. La app reemplazará el contenido por la hoja del cronograma. ' + ss.getUrl();
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp(callback, obj) {
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) return respuesta({ ok: false, error: 'Clave inválida' });
    const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) return respuesta({ ok: false, error: 'Ejecuta setup() primero' });
    const ss = SpreadsheetApp.openById(id);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    let res;
    if (data.modo === 'cargar') res = cargarCronograma(ss, data);
    else res = escribirEstados(ss, data);
    lock.releaseLock();
    return respuesta(res);
  } catch (err) {
    try { LockService.getScriptLock().releaseLock(); } catch (_) {}
    return respuesta({ ok: false, error: String(err) });
  }
}

/** Reemplaza por completo el spreadsheet con la hoja tal cual la subió el usuario. */
function cargarCronograma(ss, data) {
  const nombre = String(data.hoja || '').trim() || 'Cronograma';
  // Si ya existe una hoja con ese nombre, le cambiamos el nombre para poder borrarla.
  ss.getSheets().forEach(function (h, i) {
    if (h.getName() === nombre) {
      try { h.setName('__vieja_' + Date.now() + '_' + i); } catch (err) {}
    }
  });
  const hoja = ss.insertSheet(nombre);
  // Borra TODAS las demás: no quedan hojas viejas ni auxiliares.
  ss.getSheets().slice().forEach(function (h) {
    if (h.getSheetId() !== hoja.getSheetId()) {
      try { ss.deleteSheet(h); } catch (err) {}
    }
  });

  const valores = data.valores || [];
  if (!valores.length) return { ok: false, error: 'El Excel no trae filas' };
  const nFil = valores.length;
  let nCol = 1;
  valores.forEach(function (f) { if (f.length > nCol) nCol = f.length; });

  const vals = [], fondos = [], letras = [], negritas = [], tamanos = [], alineados = [];
  for (let i = 0; i < nFil; i++) {
    const fila = valores[i] || [];
    const rv = [], ff = [], fl = [], fn = [], fs = [], fa = [];
    for (let j = 0; j < nCol; j++) {
      const est = (data.estilos && data.estilos[i] && data.estilos[i][j]) || {};
      let v = fila[j] == null ? '' : fila[j];
      const tipo = data.tipos && data.tipos[i] && data.tipos[i][j];
      if (tipo === 'd' && v) {
        const p = String(v).slice(0, 10).split('-');
        if (p.length === 3) v = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
      }
      rv.push(v);

      const bg = String(est.bg || '').replace(/^0x/, '').replace(/^FF/, '');
      ff.push(bg.length === 6 ? '#' + bg : '#ffffff');
      let fg = String(est.fg || '').replace(/^0x/, '').replace(/^FF/, '');
      if (fg.length !== 6) fg = esOscuro(bg) ? 'ffffff' : '000000';
      fl.push('#' + fg);
      fn.push(est.b ? 'bold' : 'normal');
      fs.push(Number(est.sz || 10));
      let al = String(est.al || 'left');
      if (['left', 'center', 'right'].indexOf(al) < 0) al = 'left';
      fa.push(al);
    }
    vals.push(rv); fondos.push(ff); letras.push(fl);
    negritas.push(fn); tamanos.push(fs); alineados.push(fa);
  }

  const rango = hoja.getRange(1, 1, nFil, nCol);
  rango.setValues(vals);
  rango.setBackgrounds(fondos);
  rango.setFontColors(letras);
  rango.setFontWeights(negritas);
  rango.setFontSizes(tamanos);
  rango.setHorizontalAlignments(alineados);

  (data.merges || []).forEach(function (m) {
    try { hoja.getRange(m[0] + 1, m[1] + 1, m[2] - m[0] + 1, m[3] - m[1] + 1).merge(); } catch (err) {}
  });
  (data.cols || []).forEach(function (w, i) {
    if (w) { try { hoja.setColumnWidth(i + 1, Math.round(w)); } catch (err) {} }
  });
  (data.rows || []).forEach(function (h, i) {
    if (h) { try { hoja.setRowHeight(i + 1, Math.round(h)); } catch (err) {} }
  });
  try { hoja.setFrozenRows(Number(data.congelar || 1)); } catch (err) {}

  const aplicadas = aplicarEstados(hoja, data.filas || [], data.columnas);
  return { ok: true, hoja: nombre, filas: nFil, columnas: nCol,
    combinaciones: (data.merges || []).length, estadosAplicados: aplicadas };
}

function esOscuro(hex) {
  const h = String(hex || '');
  if (h.length !== 6) return false;
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  if (isNaN(r) || isNaN(g) || isNaN(b)) return false;
  return (r * 299 + g * 587 + b * 114) / 1000 < 140;
}

/** Escribe los resultados por equipo en la hoja indicada, sin hojas auxiliares. */
function escribirEstados(ss, data) {
  const nombre = String(data.hoja || '').trim() || 'Cronograma';
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) { hoja = ss.insertSheet(nombre); hoja.appendRow(HEADERS); }
  const aplicadas = aplicarEstados(hoja, data.filas || [], data.columnas);
  return { ok: true, hoja: hoja.getName(), modo: data.modo || 'fila',
    recibidas: (data.filas || []).length, escritas: aplicadas };
}

/** Pinta y llena SOLO los equipos realmente registrados. No toca el resto. */
function aplicarEstados(hoja, filas, columnas) {
  const registrados = (filas || []).filter(function (f) {
    return f.estado && f.estado !== 'PENDIENTE';
  });
  // Nada de registros = nada de cambios. Así el formato original se mantiene
  // intacto mientras se entra/sale de la app o corre el auto-sync.
  if (!registrados.length) return 0;

  const col = detectarColumnas(hoja, columnas);
  const filaEnc = col.filaEncabezado;
  const ultima = Math.max(hoja.getLastRow(), filaEnc);
  const seriales = hoja.getRange(filaEnc + 1, col.serial, Math.max(ultima - filaEnc, 0), 1).getValues();
  const porSerial = {};
  seriales.forEach(function (f, i) {
    const k = String(f[0] || '').trim();
    if (k) porSerial[k] = filaEnc + 1 + i;
  });
  let libre = ultima + 1;
  const ancho = Math.max(hoja.getLastColumn(), 12);
  const destinos = registrados.map(function (f) {
    const clave = String(f.serial || '').trim();
    return { f: f, fila: porSerial[clave] || libre++ };
  });

  // Se descombinan únicamente las filas que se van a escribir.
  destinos.forEach(function (d) {
    try { hoja.getRange(d.fila, 1, 1, ancho).breakApart(); } catch (err) {}
  });
  destinos.forEach(function (d) {
    const f = d.f, fila = d.fila;
    const campos = [
      [col.edificio, f.edificio], [col.piso, f.piso], [col.oficina, f.oficina],
      [col.serial, f.serial], [col.tipo, f.tipo],
      // FECHA PROGRAMADA no se toca.
      [col.ejecutada, f.fecha], [col.tecnico, f.tecnico], [col.semestre, f.semestre],
      [col.duracion, f.horasTexto || f.horas], [col.seguimiento, ''],
      [col.observaciones, f.observaciones],
    ];
    campos.forEach(function (par) {
      try { hoja.getRange(fila, par[0]).setValue(par[1] == null ? '' : par[1]); } catch (err) {}
    });
    try {
      hoja.getRange(fila, col.serial)
        .setBackground(FONDOS[f.estado] || FONDOS.PENDIENTE)
        .setFontColor('#000000');
    } catch (err) {}
  });
  return destinos.length;
}

function detectarColumnas(hoja, datos) {
  const cols = datos || {};
  const ancho = Math.max(hoja.getLastColumn(), 12);
  const maxFilas = Math.min(hoja.getLastRow(), 15);
  let filaEnc = 1;
  for (let r = 1; r <= maxFilas; r++) {
    const valores = hoja.getRange(r, 1, 1, ancho).getValues()[0];
    if (valores.some(function (v) { return /serial/i.test(String(v || '')); })) { filaEnc = r; break; }
  }
  const encontradas = {};
  hoja.getRange(filaEnc, 1, 1, ancho).getValues()[0].forEach(function (valor, i) {
    const texto = String(valor || '');
    for (const clave in PATRONES) {
      if (PATRONES[clave].test(texto) && !encontradas[clave]) encontradas[clave] = i + 1;
    }
  });
  const en = function (clave, defecto) {
    return encontradas[clave] || (typeof cols[clave] === 'number' ? cols[clave] + 1 : defecto);
  };
  return {
    filaEncabezado: filaEnc,
    edificio: en('edificio', 1), piso: en('piso', 2), oficina: en('oficina', 3),
    serial: en('serial', 4), tipo: en('tipo', 5), programada: en('programada', 6),
    ejecutada: en('ejecutada', 7), tecnico: en('tecnico', 8), semestre: en('semestre', 9),
    duracion: en('duracion', 10), seguimiento: en('seguimiento', 11), observaciones: en('observaciones', 12),
  };
}

/** Lectura para la app mediante JSONP (Apps Script no envía CORS). */
function doGet(e) {
  const p = e.parameter || {};
  const cb = /^[A-Za-z_$][\w$]*$/.test(String(p.callback || '')) ? String(p.callback) : 'prevPull';
  try {
    if (p.token !== TOKEN) return jsonp(cb, { ok: false, error: 'Clave inválida' });
    const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) return jsonp(cb, { ok: false, error: 'Ejecuta setup() primero' });
    const ss = SpreadsheetApp.openById(id);
    const nombre = String(p.hoja || '').trim();
    const hoja = ss.getSheetByName(nombre);
    if (!hoja) return jsonp(cb, { ok: false, error: 'No existe la hoja ' + nombre });

    const col = detectarColumnas(hoja, null);
    const ultima = hoja.getLastRow();
    const filas = [];
    if (ultima > col.filaEncabezado) {
      const ancho = Math.max(hoja.getLastColumn(), 12);
      const rango = hoja.getRange(col.filaEncabezado + 1, 1, ultima - col.filaEncabezado, ancho);
      const vals = rango.getValues();
      const fondos = rango.getBackgrounds();
      const inverso = {};
      for (const k in FONDOS) inverso[String(FONDOS[k]).toLowerCase()] = k;
      for (let i = 0; i < vals.length; i++) {
        const serial = String(vals[i][col.serial - 1] || '').trim();
        if (!serial) continue;
        const color = String(fondos[i][col.serial - 1] || '').toLowerCase();
        const fechaRaw = vals[i][col.ejecutada - 1];
        const fecha = fechaRaw instanceof Date
          ? fechaRaw.toISOString().slice(0, 16).replace('T', ' ')
          : String(fechaRaw || '');
        filas.push({
          serial: serial,
          estado: inverso[color] || '',
          fecha: fecha,
          tecnico: String(vals[i][col.tecnico - 1] || ''),
          semestre: String(vals[i][col.semestre - 1] || ''),
          horasTexto: String(vals[i][col.duracion - 1] || ''),
          observaciones: String(vals[i][col.observaciones - 1] || ''),
        });
      }
    }
    return jsonp(cb, { ok: true, hoja: nombre, columnas: col, filas: filas, generado: new Date().toISOString() });
  } catch (err) {
    return jsonp(cb, { ok: false, error: String(err) });
  }
}
