/* Prueba local del backend con un Google Sheets simulado. No se sube a la APK. */
(function () {
  const RESULT = { pasos: [], fallos: [], metrics: {} };
  function ok(nombre, cond, detalle) {
    RESULT.pasos.push({ nombre, ok: !!cond, detalle: detalle || '' });
    if (!cond) RESULT.fallos.push(nombre + (detalle ? ': ' + detalle : ''));
  }
  let setValueCount = 0, setValuesCalls = 0, getMergedCalls = 0, inserts = 0;

  function Sheet(nombre, filas, merges) {
    this.nombre = nombre;
    this.vals = filas.map((r) => r.slice());
    this.bg = filas.map((r) => r.map(() => '#ffffff'));
    this.font = filas.map((r) => r.map(() => '#000000'));
    this.merges = (merges || []).map((m) => ({ r0: m[0] - 1, c0: m[1] - 1, r1: m[2] - 1, c1: m[3] - 1, valor: (filas[m[0] - 1] || [])[m[1] - 1] || '' }));
    this.id = 1;
  }
  Sheet.prototype.ensure = function (r, c) {
    while (this.vals.length <= r) { this.vals.push([]); this.bg.push([]); this.font.push([]); }
    const row = this.vals[r];
    while (row.length <= c) { row.push(''); this.bg[r].push('#ffffff'); this.font[r].push('#000000'); }
  };
  Sheet.prototype.getName = function () { return this.nombre; };
  Sheet.prototype.getSheetId = function () { return this.id; };
  Sheet.prototype.getLastRow = function () { return this.vals.length; };
  Sheet.prototype.getLastColumn = function () { return Math.max(1, ...this.vals.map((r) => r.length)); };
  Sheet.prototype.getValues = function (r, c, nr, nc) {
    const out = [];
    for (let i = 0; i < nr; i++) {
      const fila = [];
      for (let j = 0; j < nc; j++) fila.push((this.vals[r - 1 + i] || [])[c - 1 + j] ?? '');
      out.push(fila);
    }
    return out;
  };
  Sheet.prototype.setValues = function (r, c, vals) {
    vals.forEach((fila, i) => fila.forEach((v, j) => { this.ensure(r - 1 + i, c - 1 + j); this.vals[r - 1 + i][c - 1 + j] = v; }));
  };
  Sheet.prototype.getRange = function (r, c, nr, nc) {
    const self = this;
    nr = nr || 1; nc = nc || 1;
    const api = {
      getValues: () => self.getValues(r, c, nr, nc),
      getValue: () => self.getValues(r, c, 1, 1)[0][0],
      setValues: (vals) => { setValuesCalls++; self.setValues(r, c, vals); return api; },
      setValue: (v) => { setValueCount++; self.setValues(r, c, [[v]]); return api; },
      getBackground: () => self.bg[r - 1][c - 1],
      getBackgrounds: () => {
        const out = [];
        for (let i = 0; i < nr; i++) { const f = []; for (let j = 0; j < nc; j++) f.push((self.bg[r - 1 + i] || [])[c - 1 + j] || '#ffffff'); out.push(f); }
        return out;
      },
      setBackground: (v) => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { self.ensure(r - 1 + i, c - 1 + j); self.bg[r - 1 + i][c - 1 + j] = v; } return api; },
      setFontColor: (v) => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { self.ensure(r - 1 + i, c - 1 + j); self.font[r - 1 + i][c - 1 + j] = v; } return api; },
      getMergedRanges: () => {
        getMergedCalls++;
        return self.merges.filter((m) => m.r0 <= r + nr - 2 && m.r1 >= r - 1 && m.c0 <= c + nc - 2 && m.c1 >= c - 1).map((m) => ({
          getRow: () => m.r0 + 1, getNumRows: () => m.r1 - m.r0 + 1,
          getColumn: () => m.c0 + 1, getNumColumns: () => m.c1 - m.c0 + 1,
          getValue: () => m.valor,
          breakApart: () => { self.merges = self.merges.filter((x) => x !== m); },
        }));
      },
      breakApart: () => { self.merges = self.merges.filter((m) => !(m.r0 <= r + nr - 2 && m.r1 >= r - 1 && m.c0 <= c + nc - 2 && m.c1 >= c - 1)); return api; },
      merge: () => { self.merges.push({ r0: r - 1, c0: c - 1, r1: r + nr - 2, c1: c + nc - 2, valor: self.vals[r - 1][c - 1] }); return api; },
    };
    return api;
  };
  Sheet.prototype.insertRowsAfter = function (fila, n) {
    inserts++;
    for (let i = 0; i < n; i++) {
      this.vals.splice(fila, 0, new Array(this.getLastColumn()).fill(''));
      this.bg.splice(fila, 0, new Array(this.getLastColumn()).fill('#ffffff'));
      this.font.splice(fila, 0, new Array(this.getLastColumn()).fill('#000000'));
    }
    this.merges.forEach((m) => { if (m.r0 >= fila) { m.r0 += n; m.r1 += n; } });
  };

  const cab = ['EDIFICIO', 'PISO', 'OFICINA', 'SERIAL UDES', 'TIPO EQUIPO', 'FECHA PROGRAMADA', 'FECHA EJECUTADA', 'TÉCNICO', 'SEMESTRE', 'TIEMPO EJECUCIÓN', 'FECHA SEGUIMIENTO', 'OBSERVACIONES'];
  const fila = (a, serial) => [a, '', a === 'A1' ? 'Of1' : 'Of2', serial, 'PC', '2026-09-15', '', '', 'B', '', '', ''];
  const hoja = new Sheet('2026-B', [cab, fila('A1', '11111111'), fila('A1', '22222222'), fila('A2', '33333333'), fila('A2', '44444444')], [[2, 1, 3, 1], [4, 1, 5, 1], [2, 2, 3, 2], [2, 3, 3, 3]]);
  const props = { SHEET_ID: 'IDDEPRUEBA1234', DIAG: '' };
  window.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getId: () => props.SHEET_ID, getSheets: () => [hoja], getSheetByName: (n) => n === hoja.nombre ? hoja : null }),
    openById: () => ({ getId: () => props.SHEET_ID, getSheets: () => [hoja], getSheetByName: (n) => n === hoja.nombre ? hoja : null }),
  };
  window.PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) };
  window.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  window.ContentService = { MimeType: { JSON: 'json', JAVASCRIPT: 'js' }, createTextOutput: (s) => ({ setMimeType() { return this; }, toString() { return s; } }) };

  fetch('apps-script/Code.gs?cb=' + Date.now()).then((r) => r.text()).then((code) => {
    new Function(code + ';window.__api={doPost,doGet,escribirEstados,cargarCronograma};')();
    const post = (obj) => JSON.parse(window.__api.doPost({ postData: { contents: JSON.stringify(obj) } }).toString());
    const base = { token: 'udes2026', hoja: '2026-B', jornada: '2026-B', modo: 'total', total: 4, fecha: '2026-11-20 10:00' };

    setValueCount = 0; setValuesCalls = 0; getMergedCalls = 0;
    let r = post(Object.assign({}, base, { filas: [] }));
    ok('sync vacío no escribe', r.ok && r.escritas === 0 && setValueCount === 0 && setValuesCalls === 0 && getMergedCalls === 0,
      JSON.stringify({ setValueCount, setValuesCalls, getMergedCalls }));

    r = post(Object.assign({}, base, { filas: [{ serial: '11111111', estado: 'OK', edificio: 'A1', piso: '', oficina: 'Of1', tipo: 'PC', fecha: '2026-11-20 10:00', tecnico: 'DANIEL', semestre: 'B', horasTexto: 'una hora', observaciones: 'sin novedades', seguimiento: '' }] }));
    ok('existente escrito', r.ok && r.escritas === 1, JSON.stringify(r));
    ok('serial con color correcto', hoja.bg[1][3].toLowerCase() === '#92d050', hoja.bg[1][3]);
    ok('serial en negro', hoja.font[1][3].toLowerCase() === '#000000', hoja.font[1][3]);
    ok('FECHA PROGRAMADA intacta', hoja.vals[1][5] === '2026-09-15', String(hoja.vals[1][5]));
    ok('FECHA SEGUIMIENTO vacía', hoja.vals[1][10] === '', String(hoja.vals[1][10]));
    ok('combinaciones separadas solo en la fila escrita',
      hoja.vals[1][0] === 'A1' && hoja.vals[2][0] === 'A1' &&
      hoja.merges.some((m) => m.r0 === 2 && m.r1 === 2 && m.c0 === 0),
      JSON.stringify({ fila2: hoja.vals[1][0], fila3: hoja.vals[2][0], merges: hoja.merges }));
    ok('pocas llamadas de escritura', setValuesCalls <= 2 && setValueCount <= 4,
      JSON.stringify({ setValuesCalls, setValueCount, getMergedCalls }));

    setValueCount = 0; setValuesCalls = 0; getMergedCalls = 0;
    const alta = { serial: '99999999', estado: 'ALTA', edificio: 'A1', piso: '', oficina: 'Of1', tipo: 'PC', fecha: '2026-11-20 11:00', tecnico: 'DANIEL', semestre: 'B', horasTexto: 'una hora', observaciones: 'nuevo', seguimiento: '' };
    r = post(Object.assign({}, base, { filas: [alta] }));
    ok('alta insertada', r.ok && inserts === 1 && hoja.vals[3][3] === '99999999', JSON.stringify({ r, inserts, row4: hoja.vals[3] }));
    ok('alta en su área (no al final)', hoja.vals[4][3] === '33333333', String(hoja.vals[4][3]));
    const antes = hoja.getLastRow();
    r = post(Object.assign({}, base, { filas: [alta] }));
    ok('alta repetida no duplica', hoja.getLastRow() === antes && inserts === 1, JSON.stringify({ filas: hoja.getLastRow(), inserts }));

    // Un cambio de color sin fecha/observaciones también debe llegar a la app.
    hoja.bg[2][3] = '#00B0F0';
    const pull = JSON.parse(window.__api.doGet({ parameter: { token: 'udes2026', hoja: '2026-B', callback: '__prevPull' } }).toString().replace(/^__prevPull\(|\);$/g, ''));
    ok('pull trae registrados y cambios de solo color',
      pull.ok && pull.filas.length === 3 &&
      pull.filas.some((f) => f.serial === '22222222' && f.estado === 'ALTA' && !f.fecha && !f.horasTexto && !f.observaciones),
      JSON.stringify(pull.filas.map((f) => [f.serial, f.estado, f.fecha, f.horasTexto, f.observaciones])));

    const mal = post(Object.assign({}, base, { hoja: 'Preventivo B 2026', filas: [alta] }));
    ok('hoja inexistente no crea hoja', mal.ok === false && /No existe la hoja/.test(mal.error || ''), JSON.stringify(mal));

    RESULT.metrics = { setValueCount, setValuesCalls, getMergedCalls, inserts, filas: hoja.getLastRow() };
    window.__RESULTADO_MOCK = RESULT;
  }).catch((e) => { RESULT.fallos.push('Error al cargar Code.gs: ' + e); window.__RESULTADO_MOCK = RESULT; });
})();
