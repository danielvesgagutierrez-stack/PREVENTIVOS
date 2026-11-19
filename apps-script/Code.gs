/**
 * Backend opcional: escribe DIRECTAMENTE en la pestaña del cronograma
 * (por ejemplo "2026-B") de este Google Sheet. No crea una hoja "Informe".
 *
 * La app envía un registro suelto cada vez que se pulsa "Guardar y seguir"
 * (modo: 'fila'). Con "Sincronizar ahora" envía la jornada completa
 * (modo: 'total'). En ambos casos se actualiza la misma pestaña.
 *
 * 1) Abre tu Google Sheet → Extensiones → Apps Script y pega este código.
 * 2) Ejecuta setup() una vez y autoriza.
 * 3) Implementar > Nueva implementación > Aplicación web >
 *    Ejecutar como: Yo > Acceso: Cualquier persona con el enlace.
 * 4) Pega la URL en la pantalla Datos de la app.
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
const COL_ESTADOS = [
  'SERIAL', 'ESTADO', 'EDIFICIO', 'PISO', 'OFICINA', 'TIPO',
  'OBSERVACIONES', 'TÉCNICO', 'HORAS', 'FECHA', 'SEMESTRE',
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
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Preventivos UDES - Informes');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  tablaEstados(ss);
  preparar(ss, 'Cambios', ['FECHA', 'SERIAL', 'TIPO', 'EDIFICIO', 'PISO', 'OFICINA', 'OBSERVACIONES', 'TÉCNICO']);
  preparar(ss, 'Resumen', ['ITEM', 'VALOR']);
  return 'Listo. La app escribirá en la pestaña del cronograma (por ejemplo 2026-B). ' + ss.getUrl();
}

function preparar(ss, nombre, headers) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = ss.insertSheet(nombre);
  if (hoja.getLastRow() === 0) hoja.appendRow(headers);
  hoja.setFrozenRows(1);
  return hoja;
}

/** Hoja oculta con el último estado de cada serial. Permite enviar de a un equipo. */
function tablaEstados(ss) {
  let hoja = ss.getSheetByName('_estados');
  if (!hoja) {
    hoja = ss.insertSheet('_estados');
    hoja.appendRow(COL_ESTADOS);
    try { hoja.hideSheet(); } catch (err) {}
  }
  return hoja;
}

/** Guarda/actualiza los seriales recibidos y devuelve TODOS los estados acumulados. */
function upsertEstados(ss, filas) {
  const hoja = tablaEstados(ss);
  const ultima = Math.max(hoja.getLastRow(), 1);
  const actuales = ultima > 1 ? hoja.getRange(2, 1, ultima - 1, COL_ESTADOS.length).getValues() : [];
  const mapa = {};
  actuales.forEach((f) => {
    const s = String(f[0] || '').trim();
    if (s) mapa[s] = f;
  });
  filas.forEach((f) => {
    const s = String(f.serial || '').trim();
    if (!s) return;
    mapa[s] = [s, f.estado || 'PENDIENTE', f.edificio || '', f.piso || '', f.oficina || '',
      f.tipo || '', f.observaciones || '', f.tecnico || '', Number(f.horas || 0),
      f.fecha || '', f.semestre || ''];
  });
  if (ultima > 1) hoja.getRange(2, 1, ultima - 1, COL_ESTADOS.length).clearContent();
  const valores = Object.keys(mapa).map((k) => mapa[k]);
  if (valores.length) hoja.getRange(2, 1, valores.length, COL_ESTADOS.length).setValues(valores);
  return valores;
}

function detectarColumnas(hoja, datos) {
  const cols = datos || {};
  const ancho = Math.max(hoja.getLastColumn(), 12);
  const maxFilas = Math.min(hoja.getLastRow(), 15);
  let filaEnc = 1;

  for (let r = 1; r <= maxFilas; r++) {
    const valores = hoja.getRange(r, 1, 1, ancho).getValues()[0];
    if (valores.some((v) => /serial/i.test(String(v || '')))) { filaEnc = r; break; }
  }

  const encontradas = {};
  hoja.getRange(filaEnc, 1, 1, ancho).getValues()[0].forEach((valor, i) => {
    const texto = String(valor || '');
    for (const clave in PATRONES) {
      if (PATRONES[clave].test(texto) && !encontradas[clave]) encontradas[clave] = i + 1;
    }
  });

  const en = (clave, defecto) =>
    encontradas[clave] || (typeof cols[clave] === 'number' ? cols[clave] + 1 : defecto);

  return {
    filaEncabezado: filaEnc,
    edificio: en('edificio', 1), piso: en('piso', 2), oficina: en('oficina', 3),
    serial: en('serial', 4), tipo: en('tipo', 5), programada: en('programada', 6),
    ejecutada: en('ejecutada', 7), tecnico: en('tecnico', 8), semestre: en('semestre', 9),
    duracion: en('duracion', 10), seguimiento: en('seguimiento', 11), observaciones: en('observaciones', 12),
  };
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.token !== TOKEN) return respuesta({ ok: false, error: 'Clave inválida' });
    const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) return respuesta({ ok: false, error: 'Ejecuta setup() primero' });
    const ss = SpreadsheetApp.openById(id);
    const filas = data.filas || [];
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    // La misma pestaña que el usuario eligió al subir el Excel (2026-B, 2026-A, ...).
    const nombreHoja = String(data.hoja || '').trim() || 'Informe';
    let hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      hoja = ss.insertSheet(nombreHoja);
      hoja.appendRow(HEADERS);
    }
    const cambios = preparar(ss, 'Cambios', ['FECHA', 'SERIAL', 'TIPO', 'EDIFICIO', 'PISO', 'OFICINA', 'OBSERVACIONES', 'TÉCNICO']);
    const resumen = preparar(ss, 'Resumen', ['ITEM', 'VALOR']);

    const col = detectarColumnas(hoja, data.columnas);
    const filaEnc = col.filaEncabezado;
    const ultima = Math.max(hoja.getLastRow(), filaEnc);

    // El cronograma tiene columnas combinadas (G2:G21, K2:K21...). Se separan
    // solo en las columnas que se escriben, para poder poner un valor por equipo.
    if (ultima > filaEnc) {
      [col.edificio, col.piso, col.oficina, col.ejecutada, col.tecnico,
        col.semestre, col.duracion, col.observaciones, col.seguimiento].forEach((c) => {
        try { hoja.getRange(filaEnc + 1, c, ultima - filaEnc, 1).breakApart(); } catch (err) {}
      });
    }

    // Mapa serial -> fila real del Sheet. Así se respeta el orden del cronograma
    // aunque el Excel y el Sheet estén ordenados distinto.
    const seriales = hoja.getRange(filaEnc + 1, col.serial, Math.max(ultima - filaEnc, 0), 1).getValues();
    const porSerial = {};
    seriales.forEach((fila, i) => {
      const clave = String(fila[0] || '').trim();
      if (clave) porSerial[clave] = filaEnc + 1 + i;
    });
    let filaLibre = ultima + 1; // equipos dados de alta que no estaban en el Sheet

    const ancho = Math.max(hoja.getLastColumn(), 12);
    const registrados = filas.filter((f) => f.estado && f.estado !== 'PENDIENTE');
    registrados.forEach((f) => {
      const clave = String(f.serial || '').trim();
      const fila = porSerial[clave] || filaLibre++;
      const valores = hoja.getRange(fila, 1, 1, ancho).getValues()[0];

      valores[col.edificio - 1] = f.edificio || '';
      valores[col.piso - 1] = f.piso || '';
      valores[col.oficina - 1] = f.oficina || '';
      valores[col.serial - 1] = f.serial || '';
      valores[col.tipo - 1] = f.tipo || '';
      // FECHA PROGRAMADA se deja como estaba en el Sheet.
      valores[col.ejecutada - 1] = f.fecha || '';
      valores[col.tecnico - 1] = f.tecnico || '';
      valores[col.semestre - 1] = f.semestre || '';
      valores[col.duracion - 1] = f.horasTexto || f.horas || '';
      valores[col.seguimiento - 1] = ''; // la llena otra persona
      valores[col.observaciones - 1] = f.observaciones || '';

      hoja.getRange(fila, 1, 1, ancho).setValues([valores]);
      // Solo cambia el fondo; el número de inventario queda en negro.
      hoja.getRange(fila, col.serial)
        .setBackground(FONDOS[f.estado] || FONDOS.PENDIENTE)
        .setFontColor('#000000');
    });

    // Acumulamos TODOS los estados (de a un equipo o de la jornada completa).
    const todos = upsertEstados(ss, filas);

    // Cambios: altas, reasignaciones y no encontrados (se recalcula con cada envío).
    const camb = todos.filter((v) => v[1] === 'ALTA' || v[1] === 'AUSENTE').map((v) => [
      v[9] || new Date().toISOString(), v[0],
      v[1] === 'ALTA' ? 'OTRA DEPENDENCIA / NUEVO' : 'NO ENCONTRADO',
      v[2], v[3], v[4], v[6], v[7],
    ]);
    if (cambios.getLastRow() > 1) cambios.getRange(2, 1, cambios.getLastRow() - 1, 8).clearContent();
    if (camb.length) cambios.getRange(2, 1, camb.length, 8).setValues(camb);

    // Resumen: se reconstruye desde el acumulado, no se pierde con envíos parciales.
    const conteo = { OK: 0, NOVEDAD: 0, AUSENTE: 0, ALTA: 0, PENDIENTE: 0 };
    let horas = 0;
    todos.forEach((v) => { conteo[v[1]] = (conteo[v[1]] || 0) + 1; horas += Number(v[8] || 0); });
    const registradosUnicos = todos.filter((v) => v[1] !== 'PENDIENTE').length;
    const total = Number(data.total || registradosUnicos);
    const filasRes = [
      ['Hoja', nombreHoja], ['Jornada', data.jornada || ''], ['Técnico', data.tecnico || ''],
      ['Generado', data.fecha || ''], ['OK', conteo.OK], ['Novedad', conteo.NOVEDAD],
      ['No encontrado', conteo.AUSENTE], ['Cambio de dependencia / nuevo', conteo.ALTA],
      ['Pendientes', Math.max(total - registradosUnicos, 0)], ['Total horas', horas],
      ['Fecha ejecución', 'Automática'], ['Fecha seguimiento', 'En blanco'],
    ];
    if (resumen.getLastRow() > 1) resumen.getRange(2, 1, resumen.getLastRow() - 1, 2).clearContent();
    resumen.getRange(2, 1, filasRes.length, 2).setValues(filasRes);

    lock.releaseLock();
    return respuesta({ ok: true, hoja: nombreHoja, modo: data.modo || 'fila',
      recibidas: filas.length, escritas: registrados.length, acumuladas: todos.length });
  } catch (err) {
    try { LockService.getScriptLock().releaseLock(); } catch (_) {}
    return respuesta({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutput('<h2>Preventivos UDES</h2><p>Backend activo. Configura la URL en la app.</p>');
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
