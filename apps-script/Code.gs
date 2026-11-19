/**
 * Backend opcional: recibe lo que registra la PWA y deja una hoja "Informe"
 * con el mismo formato del cronograma (columnas A-L y colores por estado).
 *
 * 1) Pegar en https://script.google.com
 * 2) Ejecutar setup() una vez y autorizar.
 * 3) Implementar > Nueva implementación > Aplicación web >
 *    Ejecutar como: Yo > Acceso: Cualquier persona con el enlace.
 * 4) Pegar la URL en la pantalla Datos de la app.
 */
const TOKEN = 'udes2026';
const COLORES = {
  'FF92D050': ['#92D050', '#15803d'],
  'FFFFC000': ['#FFC000', '#9a3412'],
  'FFFF0000': ['#FF0000', '#b91c1c'],
  'FF00B0F0': ['#00B0F0', '#1d4ed8'],
};
const HEADERS = [
  'EDIFICIO', 'PISO', 'OFICINA', 'SERIAL UDES', 'TIPO EQUIPO', 'FECHA PROGRAMADA',
  'FECHA EJECUTADA', 'TÉCNICO', 'SEMESTRE', 'TIEMPO EJECUCIÓN', 'FECHA SEGUIMIENTO', 'OBSERVACIONES',
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Preventivos UDES - Informes');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  preparar(ss, 'Informe', HEADERS);
  preparar(ss, 'Cambios', ['FECHA', 'SERIAL', 'TIPO', 'EDIFICIO', 'PISO', 'OFICINA', 'OBSERVACIONES', 'TÉCNICO']);
  preparar(ss, 'Resumen', ['ITEM', 'VALOR']);
  return 'Listo. Hoja: ' + ss.getUrl();
}

function preparar(ss, nombre, headers) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = ss.insertSheet(nombre);
  if (hoja.getLastRow() === 0) hoja.appendRow(headers);
  hoja.setFrozenRows(1);
  return hoja;
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

    const informe = ss.getSheetByName('Informe');
    const cambios = ss.getSheetByName('Cambios');
    const resumen = ss.getSheetByName('Resumen');
    const cols = data.columnas || {};
    const en = (k, def) => (typeof cols[k] === 'number' ? cols[k] + 1 : def + 1);
    const cEdificio = en('edificio', 0), cPiso = en('piso', 1), cOficina = en('oficina', 2);
    const cSerial = en('serial', 3), cTipo = en('tipo', 4), cProgramada = en('programada', 5);
    const cEjecutada = en('ejecutada', 6), cTecnico = en('tecnico', 7), cSemestre = en('semestre', 8);
    const cDuracion = en('duracion', 9), cSeguimiento = en('seguimiento', 10), cObs = en('observaciones', 11);

    // Informe: una fila por equipo, con colores.
    const ultima = Math.max(informe.getLastRow(), 1);
    if (ultima > 1) informe.getRange(2, 1, ultima - 1, 12).clear();
    const valores = [], fondos = [], textos = [], negritas = [];
    filas.forEach((f) => {
      const fila = new Array(12).fill('');
      fila[cEdificio - 1] = f.edificio || '';
      fila[cPiso - 1] = f.piso || '';
      fila[cOficina - 1] = f.oficina || '';
      fila[cSerial - 1] = f.serial || '';
      fila[cTipo - 1] = f.tipo || '';
      fila[cProgramada - 1] = f.programada || '';
      fila[cEjecutada - 1] = f.fecha || '';
      fila[cTecnico - 1] = f.tecnico || '';
      fila[cSemestre - 1] = f.semestre || '';
      fila[cDuracion - 1] = f.horasTexto || f.horas || '';
      fila[cSeguimiento - 1] = '';
      fila[cObs - 1] = f.observaciones || '';
      valores.push(fila);
      const c = COLORES[f.color] || ['#e9eef4', '#64748b'];
      fondos.push(new Array(12).fill('#ffffff').map((v, i) => (i === cSerial - 1 ? c[0] : v)));
      textos.push(new Array(12).fill('#1c2733').map((v, i) => (i === cSerial - 1 ? c[1] : v)));
      negritas.push(new Array(12).fill(false).map((v, i) => i === cSerial - 1));
    });
    if (valores.length) {
      informe.getRange(2, 1, valores.length, 12).setValues(valores);
      informe.getRange(2, 1, valores.length, 12).setBackgrounds(fondos);
      informe.getRange(2, 1, valores.length, 12).setFontColors(textos);
      informe.getRange(2, 1, valores.length, 12).setFontWeights(negritas);
      informe.getRange(2, 1, valores.length, 12).setNumberFormat('@');
    }

    // Cambios: altas, reasignaciones y no encontrados.
    const camb = filas.filter((f) => ['ALTA', 'AUSENTE'].indexOf(f.estado) !== -1).map((f) => [
      data.fecha || new Date().toISOString(), f.serial,
      f.estado === 'ALTA' ? 'OTRA DEPENDENCIA / NUEVO' : 'NO ENCONTRADO',
      f.edificio, f.piso, f.oficina, f.observaciones, f.tecnico,
    ]);
    if (cambios.getLastRow() > 1) cambios.getRange(2, 1, cambios.getLastRow() - 1, 8).clearContent();
    if (camb.length) cambios.getRange(2, 1, camb.length, camb[0].length).setValues(camb);

    // Resumen.
    const conteo = { OK: 0, NOVEDAD: 0, AUSENTE: 0, ALTA: 0, PENDIENTE: 0 };
    let horas = 0;
    filas.forEach((f) => { conteo[f.estado] = (conteo[f.estado] || 0) + 1; horas += Number(f.horas || 0); });
    const filasRes = [
      ['Hoja', data.hoja || ''], ['Jornada', data.jornada || ''], ['Técnico', data.tecnico || ''],
      ['Generado', data.fecha || ''], ['OK', conteo.OK], ['Novedad', conteo.NOVEDAD],
      ['No encontrado', conteo.AUSENTE], ['Cambio de dependencia / nuevo', conteo.ALTA],
      ['Pendientes', conteo.PENDIENTE], ['Total horas', horas],
      ['Fecha ejecución', 'Automática'], ['Fecha seguimiento', 'En blanco'],
    ];
    if (resumen.getLastRow() > 1) resumen.getRange(2, 1, resumen.getLastRow() - 1, 2).clearContent();
    resumen.getRange(2, 1, filasRes.length, 2).setValues(filasRes);

    lock.releaseLock();
    return respuesta({ ok: true, recibidas: filas.length });
  } catch (err) {
    try { LockService.getScriptLock().releaseLock(); } catch (_) {}
    return respuesta({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutput('<h2>Preventivos UDES</h2><p>Backend activo.</p>');
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
