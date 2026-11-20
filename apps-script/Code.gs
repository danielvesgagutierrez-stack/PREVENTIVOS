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
const VERSION = '2026-11-20e'; // se muestra en Datos para confirmar que está desplegado
let diagEscritura = { fallos: [] }; // se guarda en doPost y se lee desde la app
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

/** Guarda el resultado del último envío para que la app pueda leerlo por JSONP. */
function guardarDiag(obj) {
  try {
    PropertiesService.getScriptProperties().setProperty('DIAG', JSON.stringify(
      Object.assign({ hora: new Date().toISOString(), version: VERSION }, obj)));
  } catch (err) {}
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
    try {
      if (data.modo === 'cargar') res = cargarCronograma(ss, data);
      else res = escribirEstados(ss, data);
    } finally {
      lock.releaseLock();
    }
    guardarDiag(Object.assign({ ok: res.ok !== false, modo: data.modo,
      hoja: data.hoja || '', id: String(id).slice(0, 6) + '…' + String(id).slice(-4),
      recibidas: (data.filas || []).length }, res, { diag: diagEscritura }));
    return respuesta(res);
  } catch (err) {
    try { LockService.getScriptLock().releaseLock(); } catch (_) {}
    guardarDiag({ ok: false, error: String(err), diag: diagEscritura });
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

  const columnas = data.columnas || {};
  const colSerial = Number(columnas.serial) + 1;                       // 1-based
  const primeraFilaDatos = Number(columnas.headerRow || 0) + 1;        // índice 0-based
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

      const bg = hex6(est.bg);
      ff.push(bg ? '#' + bg : '#ffffff');
      let fg = hex6(est.fg);
      if (!fg) fg = esOscuro(bg) ? 'ffffff' : '000000';
      // El número de inventario siempre en negro; solo cambia el fondo de la celda.
      fl.push((j + 1 === colSerial && i >= primeraFilaDatos) ? '#000000' : '#' + fg);
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

/** Clave de área sin acentos para comparar Excel y Sheet aunque difieran en tildes. */
function normZona(valor) {
  return String(valor || '').trim().toUpperCase()
    .replace(/[ÁÀÄÅ]/g, 'A').replace(/[ÉÈË]/g, 'E').replace(/[ÍÌÏ]/g, 'I')
    .replace(/[ÓÒÖ]/g, 'O').replace(/[ÚÙÜ]/g, 'U').replace(/Ñ/g, 'N');
}

/** Normaliza 0xFF92D050 / FF92D050 / 92D050 a 92D050 (o ''). */
function hex6(valor) {
  let h = String(valor || '').replace(/^0x/i, '').toUpperCase();
  if (h.length === 8 && h.slice(0, 2) === 'FF') h = h.slice(2);
  return h.length === 6 ? h : '';
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
  const nombre = String(data.hoja || '').trim();
  const hoja = ss.getSheetByName(nombre);
  // NUNCA crear hojas aquí: si el nombre no coincide, la app escribiría en una
  // hoja nueva y parecería que no pasa nada. Mejor avisar con el nombre exacto.
  if (!hoja) return { ok: false, error: 'No existe la hoja "' + nombre + '". Sube el cronograma o corrige el nombre en Datos.',
    hojas: ss.getSheets().map(function (h) { return h.getName(); }) };
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
  diagEscritura = { filas: registrados.length, existentes: 0, nuevos: 0, fallos: [] };
  if (!registrados.length) return 0;

  const col = detectarColumnas(hoja, columnas);
  const filaEnc = col.filaEncabezado;
  const ancho = Math.max(hoja.getLastColumn(), 12);
  const anchoArea = Math.max(col.edificio, col.piso, col.oficina);

  /**
   * Serial -> fila y área -> ÚLTIMA fila real de esa área.
   * Las columnas combinadas (A2:A50, C2:C20...) solo guardan el valor en la
   * primera fila; se rellena hacia abajo para conocer el área de CADA fila.
   * Se recalcula antes de cada inserción porque insertRowsAfter corre las filas.
   */
  function mapaActual() {
    const ultima = Math.max(hoja.getLastRow(), filaEnc);
    const n = Math.max(ultima - filaEnc, 0);
    const todo = n ? hoja.getRange(filaEnc + 1, 1, n, ancho).getValues() : [];
    const porSerial = {}, ultimaZona = {}, vacias = {};
    let ed = '', pi = '', of = '';
    for (let i = 0; i < n; i++) {
      const fila = filaEnc + 1 + i;
      const vE = normZona(todo[i][col.edificio - 1]);
      const vP = normZona(todo[i][col.piso - 1]);
      const vO = normZona(todo[i][col.oficina - 1]);
      if (vE) ed = vE;
      if (vP) pi = vP;
      if (vO) of = vO;
      const zona = [ed, pi, of].join('|');
      if (zona.replace(/\|/g, '')) ultimaZona[zona] = fila;
      const k = String(todo[i][col.serial - 1] || '').replace(/\D/g, '');
      if (k) porSerial[k] = fila;
      else if (todo[i].every(function (v) { return !String(v || '').trim(); })) {
        // Fila que un sync anterior dejó en blanco: se reutiliza en vez de
        // insertar otra (evita el "doble celda" al guardar de nuevo).
        if (!vacias[zona]) vacias[zona] = [];
        vacias[zona].push(fila);
      }
    }
    return { porSerial: porSerial, ultimaZona: ultimaZona, vacias: vacias, ultima: ultima };
  }
  const zonaDe = function (f) {
    return [f.edificio, f.piso, f.oficina].map(normZona).join('|');
  };
  const claveSerial = function (v) { return String(v || '').replace(/\D/g, ''); };
  const COLS_ESCRITAS = [col.edificio, col.piso, col.oficina, col.serial, col.tipo,
    col.ejecutada, col.tecnico, col.semestre, col.duracion, col.seguimiento, col.observaciones];

  /**
   * breakApart() solo funciona si el rango cubre la combinación COMPLETA.
   * Por eso se piden las combinaciones reales con getMergedRanges() y se parte
   * únicamente la que toca la fila que vamos a escribir; después se vuelve a
   * unir el trozo de arriba y el de abajo para conservar el formato.
   */
  function romperCombinadas(fila) {
    try {
      const n = Math.max(hoja.getLastRow() - filaEnc, 1);
      const lista = hoja.getRange(filaEnc + 1, 1, n, ancho).getMergedRanges();
      lista.forEach(function (mr) {
        const ini = mr.getRow(), fin = ini + mr.getNumRows() - 1;
        if (fila < ini || fila > fin) return;
        const c0 = mr.getColumn(), nc = mr.getNumColumns();
        let valor = '';
        try { valor = mr.getValue(); } catch (err) {}
        try { mr.breakApart(); } catch (err) {
          diagEscritura.fallos.push('unir fila ' + fila + ': ' + String(err));
        }
        try { if (fila > ini) hoja.getRange(ini, c0, fila - ini, nc).merge(); } catch (err) {}
        try {
          if (fila < fin) {
            const abajo = hoja.getRange(fila + 1, c0, fin - fila, nc);
            abajo.setValue(valor); abajo.merge();
          }
        } catch (err) {}
      });
    } catch (err) {
      // Compatibilidad: si getMergedRanges no está, se parten columnas enteras.
      COLS_ESCRITAS.forEach(function (c) {
        try { hoja.getRange(filaEnc + 1, c, Math.max(hoja.getLastRow() - filaEnc, 1), 1).breakApart(); } catch (e) {}
      });
    }
  }

  function escribir(f, fila) {
    romperCombinadas(fila);
    const campos = [
      [col.edificio, f.edificio], [col.piso, f.piso], [col.oficina, f.oficina],
      [col.serial, f.serial], [col.tipo, f.tipo],
      // FECHA PROGRAMADA no se toca.
      [col.ejecutada, f.fecha], [col.tecnico, f.tecnico], [col.semestre, f.semestre],
      [col.duracion, f.horasTexto || f.horas], [col.seguimiento, ''],
      [col.observaciones, f.observaciones],
    ];
    campos.forEach(function (par) {
      try { hoja.getRange(fila, par[0]).setValue(par[1] == null ? '' : par[1]); }
      catch (err) { diagEscritura.fallos.push('fila ' + fila + ' col ' + par[0] + ': ' + String(err)); }
    });
    const fondo = FONDOS[f.estado] || FONDOS.PENDIENTE;
    try {
      hoja.getRange(fila, col.serial).setBackground(fondo).setFontColor('#000000');
    } catch (err) { diagEscritura.fallos.push('fila ' + fila + ' color: ' + String(err)); }

    // Verificación real: si el serial o el color no quedaron, partimos todas
    // las columnas de la zona y repetimos. Así nunca queda una fila en blanco.
    const eco = claveSerial(hoja.getRange(fila, col.serial).getValue());
    const ecoBg = String(hoja.getRange(fila, col.serial).getBackground() || '').toLowerCase();
    if (eco !== claveSerial(f.serial) || ecoBg !== fondo.toLowerCase()) {
      const n = Math.max(hoja.getLastRow() - filaEnc, 1);
      COLS_ESCRITAS.forEach(function (c) {
        try { hoja.getRange(filaEnc + 1, c, n, 1).breakApart(); } catch (err) {}
      });
      campos.forEach(function (par) {
        try { hoja.getRange(fila, par[0]).setValue(par[1] == null ? '' : par[1]); } catch (err) {}
      });
      try { hoja.getRange(fila, col.serial).setBackground(fondo).setFontColor('#000000'); } catch (err) {}
    }
  }

  // 1) Existentes: sus filas NO cambian. Se escriben ANTES de insertar nada,
  //    para que el color y los datos caigan en el serial correcto.
  const mapa = mapaActual();
  const existentes = [], nuevos = [];
  registrados.forEach(function (f) {
    const k = claveSerial(f.serial);
    if (mapa.porSerial[k]) existentes.push({ f: f, fila: mapa.porSerial[k] });
    else nuevos.push(f);
  });
  diagEscritura.existentes = existentes.length;
  diagEscritura.nuevos = nuevos.length;
  existentes.forEach(function (d) { escribir(d.f, d.fila); });

  // 2) Nuevos: primero reutilizan una fila vacía que haya dejado un sync
  //    anterior; si no la hay, se insertan JUSTO DESPUÉS de la última fila de
  //    su área (junto a Antiguo Masira, no al final). Se recalcula el mapa en
  //    cada inserción para no escribir nunca sobre índices viejos.
  nuevos.forEach(function (f) {
    const m = mapaActual();
    const zona = zonaDe(f);
    const libre = (m.vacias[zona] || [])[0];
    if (libre) { escribir(f, libre); return; }
    const base = m.ultimaZona[zona] || m.ultima;
    hoja.insertRowsAfter(base, 1);
    escribir(f, base + 1);
  });

  return registrados.length;
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
    let ultimo = null;
    try { ultimo = JSON.parse(PropertiesService.getScriptProperties().getProperty('DIAG') || 'null'); } catch (err) {}
    return jsonp(cb, { ok: true, version: VERSION, ultimo: ultimo,
      hoja: nombre, columnas: col, filas: filas, generado: new Date().toISOString() });
  } catch (err) {
    return jsonp(cb, { ok: false, error: String(err) });
  }
}
