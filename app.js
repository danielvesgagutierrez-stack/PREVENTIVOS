/* Preventivos UDES - PWA offline-first
 * El usuario sube el Excel, elige la hoja del mantenimiento y registra por equipo:
 * OK | NOVEDAD | AUSENTE | ALTA (nuevo o reasignado a otra dependencia).
 * El informe final se genera con el formato del archivo original.
 */
(() => {
  'use strict';

  const K = {
    data: 'prev_dataset_v3',
    progress: 'prev_progress_v3',
    config: 'prev_config_v3',
    jornada: 'prev_jornada_v3',
    archivo: 'prev_archivo_v3',
    pendiente: 'prev_pendiente_v1', // seriales que no se pudieron enviar (sin internet)
    sync: 'prev_sync_v1',           // resultado de la última sincronización
    cargar: 'prev_cargar_v1',       // cronograma pendiente de subir al Sheet
  };

  const ESTADOS = {
    OK:        { label: 'Hecho',                    desc: 'Sin novedades',                  argb: 'FF92D050' },
    NOVEDAD:   { label: 'Novedad',                  desc: 'Pantalla partida, etc.',         argb: 'FFFFC000' },
    AUSENTE:   { label: 'No se encuentra',          desc: 'Ya no está en la dependencia',   argb: 'FFFF0000' },
    ALTA:      { label: 'Otra dependencia / nuevo', desc: 'Se reasignó o es equipo nuevo',  argb: 'FF00B0F0' },
    PENDIENTE: { label: 'Pendiente',                desc: 'Sin registrar',                  argb: 'FFE9EEF4' },
  };
  const ORDEN = ['OK', 'NOVEDAD', 'AUSENTE', 'ALTA'];
  const COLUMNAS_POR_DEFECTO = {
    edificio: 0, piso: 1, oficina: 2, serial: 3, tipo: 4, programada: 5,
    ejecutada: 6, tecnico: 7, semestre: 8, duracion: 9, seguimiento: 10,
    observaciones: 11, headerRow: 0,
  };

  const state = {
    data: null, jornadaId: null, progress: {}, config: {}, archivoMeta: null,
    ruta: '#/inicio', areaActual: null, seleccion: null, seleccionSerial: null,
    horas: '1 hora', filtroEdificio: '', filtroTexto: '', hojas: [],
    ultimoScan: { texto: '', hora: 0 }, filtroTimer: null,
  };

  const $ = (sel) => document.querySelector(sel);
  const view = $('#view');

  /* ---------------- utilidades ---------------- */

  function esc(v) {
    return String(v ?? '').replace(/[&<>"']/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function norm(v) { return String(v ?? '').replace(/\s+/g, ' ').trim(); }
  function clave(v) {
    return norm(v).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^A-Z0-9]+/g, '-').replace(/^-|-$/g, '');
  }
  function digitos(v) { return String(v ?? '').replace(/\D/g, ''); }
  function hoy() { return new Date().toISOString().slice(0, 10); }
  function hora() { return new Date().toISOString().slice(0, 16).replace('T', ' '); }
  function leer(k, d) { try { return JSON.parse(localStorage.getItem(k)) ?? d; } catch (_) { return d; } }
  function escribir(k, v) {
    try { localStorage.setItem(k, JSON.stringify(v)); }
    catch (_) { toast('Sin espacio en el teléfono'); }
  }
  let toastTimer;
  function toast(msg) {
    const t = $('#toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 2800);
  }
  function vibrar(p = 90) { try { navigator.vibrate?.(p); } catch (_) {} }
  function go(hash) { location.hash = hash; }
  function chip(e) { return `<span class="chip ${e}">${ESTADOS[e]?.label || e}</span>`; }
  function horasTexto(h) {
    const n = Number(h);
    if (!n) return 'una hora';
    if (n === 0.5) return '30 minutos';
    if (n === 1) return 'una hora';
    if (n === 1.5) return 'una hora y media';
    if (n === 2) return 'dos horas';
    return n + ' horas';
  }
  /** Convierte texto libre ("media hora", "1,5 horas", etc.) a número. */
  function horasDeTexto(txt) {
    const t = norm(txt);
    if (!t) return 0;
    if (/media hora|30 minutos/i.test(t)) return 0.5;
    if (/hora y media/i.test(t)) return 1.5;
    if (/dos horas/i.test(t)) return 2;
    const m = t.match(/(\d+(?:[.,]\d+)?)/);
    return m ? Number(m[1].replace(',', '.')) : 0;
  }

  function descargar(datos, nombre, mime) {
    const url = URL.createObjectURL(new Blob([datos], { type: mime }));
    const a = document.createElement('a'); a.href = url; a.download = nombre;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 8000);
  }

  /* ---------------- IndexedDB: fotos y archivo original ---------------- */

  function db() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('prev_datos', 2);
      r.onupgradeneeded = () => {
        const d = r.result;
        if (!d.objectStoreNames.contains('fotos')) d.createObjectStore('fotos');
        if (!d.objectStoreNames.contains('archivos')) d.createObjectStore('archivos');
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function idbPoner(tienda, clave, valor) {
    const d = await db();
    return new Promise((res, rej) => {
      const tx = d.transaction(tienda, 'readwrite');
      tx.objectStore(tienda).put(valor, clave);
      tx.oncomplete = res; tx.onerror = () => rej(tx.error);
    });
  }
  async function idbObtener(tienda, clave) {
    try {
      const d = await db();
      return await new Promise((res) => {
        const q = d.transaction(tienda).objectStore(tienda).get(clave);
        q.onsuccess = () => res(q.result || null); q.onerror = () => res(null);
      });
    } catch (_) { return null; }
  }
  const fotoGuardar = (s, url) => idbPoner('fotos', state.jornadaId + '|' + s, url);
  const fotoLeer = (s) => idbObtener('fotos', state.jornadaId + '|' + s);
  const archivoGuardar = (blob) => idbPoner('archivos', 'original', blob);
  const archivoLeer = () => idbObtener('archivos', 'original');

  /* ---------------- estado ---------------- */

  function guardarProgreso() { escribir(K.progress, state.progress); }
  function guardarDataset() { escribir(K.data, state.data); }
  function jornada() { return state.data?.jornadas?.[0]; }
  function nombreHoja() { return state.archivoMeta?.hoja || state.data?.hoja || jornada()?.nombre || 'Informe'; }
  function items() { return jornada()?.items || []; }
  function areas() { return state.data?.areas || []; }
  function columnas() { return state.data?.columnas || COLUMNAS_POR_DEFECTO; }
  function progreso(jid, serial) { return state.progress[jid]?.[serial] || null; }
  // Solo cuenta lo registrado de verdad (Guardar y seguir / cambio traído del Sheet).
  // Los colores que vienen en el Excel son el cronograma, no un registro real.
  function registradoDe(item) {
    const p = progreso(state.jornadaId, item.serial);
    return p && p.registrado ? p.estado : null;
  }
  function estadoDe(item) { return registradoDe(item) || 'PENDIENTE'; }
  function novedadDe(item) { return progreso(state.jornadaId, item.serial)?.novedad || item.novedad || ''; }
  function horasDe(item) {
    const p = progreso(state.jornadaId, item.serial);
    const porDefecto = Number(String(item.duracion || '').replace(/\D/g, '')) || 1;
    return p?.horas ?? porDefecto;
  }

  function stats() {
    const j = jornada(); const c = { OK: 0, NOVEDAD: 0, AUSENTE: 0, ALTA: 0, PENDIENTE: 0, HORAS: 0 };
    for (const it of j?.items || []) {
      const p = progreso(j.id, it.serial) || {};
      // Un color del Excel NO es un registro: si no se grabó, cuenta como pendiente.
      const e = p.registrado ? p.estado : 'PENDIENTE';
      c[e] = (c[e] || 0) + 1;
      if (e !== 'PENDIENTE') c.HORAS += Number(p.horas || horasDe(it) || 0);
    }
    c.TOTAL = j?.items.length || 0;
    c.HECHOS = c.OK + c.NOVEDAD + c.AUSENTE + c.ALTA;
    return c;
  }
  function statsHtml(s) {
    return `<div class="stats">
      <div class="stat"><b>${s.TOTAL}</b><span>Total</span></div>
      <div class="stat OK"><b>${s.OK}</b><span>Hecho</span></div>
      <div class="stat NOVEDAD"><b>${s.NOVEDAD}</b><span>Novedad</span></div>
      <div class="stat AUSENTE"><b>${s.AUSENTE}</b><span>Ausente</span></div>
      <div class="stat ALTA"><b>${s.ALTA}</b><span>Cambio</span></div>
    </div>`;
  }

  /* ---------------- carga ---------------- */

  async function cargar() {
    state.config = leer(K.config, { url: '', token: 'udes2026', tecnico: '' });
    state.progress = leer(K.progress, {});
    state.archivoMeta = leer(K.archivo, null);
    const guardado = leer(K.data, null);
    if (guardado?.jornadas?.length) state.data = guardado;
    else {
      try {
        const r = await fetch('data/equipos.json', { cache: 'no-cache' });
        state.data = r.ok ? await r.json() : null;
        if (state.data) guardarDataset();
      } catch (_) { state.data = null; }
    }
    if (state.data?.jornadas?.length) state.jornadaId = state.data.jornadas[0].id;
  }

  /* ---------------- render ---------------- */

  function rutaActual() { return (location.hash || '#/inicio').split('?')[0]; }
  function partes() { return rutaActual().replace(/^#\/?/, '').split('/'); }
  function areaNombre(id) {
    const a = areas().find((x) => x.id === id);
    return a ? [a.oficina, a.piso, a.edificio].filter(Boolean).join(' · ') : (id || '—');
  }

  function render() {
    detenerEscanner();
    state.ruta = rutaActual();
    const p = partes();
    const vistas = {
      inicio: viewInicio, areas: viewAreas, area: viewArea, escanear: viewEscanear,
      equipo: viewEquipo, alta: viewAlta, datos: viewDatos, ayuda: viewAyuda,
      etiquetas: viewEtiquetas, 'elegir-hoja': viewElegirHoja,
    };
    const fn = vistas[p[0]] || viewInicio;
    $('#topbar').hidden = p[0] === 'escanear';
    const j = jornada();
    document.getElementById('titulo').textContent = {
      inicio: 'Preventivos', areas: 'Áreas y dependencias', area: p[1] ? areaNombre(decodeURIComponent(p[1])) : 'Área',
      escanear: 'Escanear', equipo: 'Equipo', alta: 'Dar de alta', datos: 'Datos e informe',
      ayuda: 'Cómo se usa', etiquetas: 'Etiquetas', 'elegir-hoja': 'Elegir hoja',
    }[p[0]] || 'Preventivos';
    document.getElementById('subtitulo').textContent = p[0] === 'escanear'
      ? 'Apunta la cámara al código de barras'
      : j ? `${j.nombre} · ${j.tecnico || 'sin técnico'}` : 'Sin cronograma cargado';
    $('#btnBack').hidden = ['inicio', 'areas', 'escanear', 'datos', 'elegir-hoja'].includes(p[0]);
    document.querySelectorAll('#bottomnav button').forEach((b) => {
      b.classList.toggle('active', b.dataset.nav === p[0] || (p[0] === 'equipo' && b.dataset.nav === 'escanear'));
    });
    view.innerHTML = fn(p);
    window.scrollTo(0, 0);
    if (p[0] === 'escanear') setTimeout(iniciarEscanner, 40);
    if (p[0] === 'equipo') {
      fotoLeer(decodeURIComponent(p[1] || '')).then((url) => {
        const d = $('#fotoActual');
        if (d && url) d.innerHTML = `<img src="${url}" alt="Foto" style="max-width:100%;border-radius:10px">`;
      });
    }
  }

  /* ---------------- vistas ---------------- */

  function sinDatos() {
    return `<div class="card vacio"><b>No hay cronograma cargado</b>
      Sube el Excel del mantenimiento desde Inicio o Datos.</div>`;
  }

  function viewInicio() {
    if (!state.data) return sinDatos();
    const s = stats(); const j = jornada();
    const totalHoras = Number(s.HORAS.toFixed(1));
    const porEdificio = {};
    for (const it of items()) {
      const e = it.edificio || 'Sin edificio';
      porEdificio[e] ||= { total: 0, hechos: 0 };
      porEdificio[e].total++;
      if (estadoDe(it) !== 'PENDIENTE') porEdificio[e].hechos++;
    }
    const recientes = Object.entries(state.progress[state.jornadaId] || {})
      .sort((a, b) => String(b[1].fecha || '').localeCompare(String(a[1].fecha || ''))).slice(0, 5);
    return `
      <div class="card">
        <h2>Hoja del mantenimiento</h2>
        <div class="row spread"><div class="grow"><b>${esc(j.nombre)}</b>
          <div class="small muted">${esc(state.archivoMeta?.nombre || state.data.fuente || '')} · ${esc(j.tecnico || 'Sin técnico')}</div></div>
          <button class="btn sm" data-act="subir">Cambiar</button></div>
        <div class="mt">${statsHtml(s)}</div>
        <div class="small muted center mt">${s.HECHOS}/${s.TOTAL} registrados · ${Math.round((s.HECHOS / (s.TOTAL || 1)) * 100)}% · ${totalHoras} h ejecutadas</div>
      </div>
      <div class="btn-grid mb">
        <button class="btn primary" data-act="go" data-hash="#/escanear">◎ Escanear equipo</button>
        <button class="btn" data-act="go" data-hash="#/areas">▤ Ver por área</button>
        <button class="btn" data-act="subir">⇧ Subir cronograma</button>
        <button class="btn primary" data-act="informe">↧ Generar informe final</button>
      </div>
      <div class="card"><h2>Avance por edificio</h2>
        ${Object.entries(porEdificio).sort((a, b) => a[0].localeCompare(b[0])).map(([e, d]) => `
          <div class="item"><div class="item-main"><strong>${esc(e)}</strong>
            <div class="bar mt"><i class="b-OK" style="width:${(d.hechos / d.total) * 100}%"></i></div></div>
            <span class="small muted">${d.hechos}/${d.total}</span></div>`).join('')}
      </div>
      ${recientes.length ? `<div class="card"><h2>Últimos registros</h2><div class="lista">
        ${recientes.map(([serial, r]) => `<div class="item"><div class="item-main">
          <strong class="mono">${esc(serial)}</strong>
          <small>${esc(ESTADOS[r.estado]?.label || r.estado)} · ${esc(r.horasTexto || '')} · ${esc(r.novedad || 'sin novedad')} · ${esc(r.fecha || '')}</small></div>
          ${chip(r.estado)}</div>`).join('')}</div></div>` : ''}`;
  }

  function viewElegirHoja() {
    if (!state.hojas.length) return `<div class="card vacio">No se encontraron hojas con equipos.</div>`;
    return `<div class="card">
      <h2>¿Cuál hoja es la de este mantenimiento?</h2>
      <p class="small muted">El archivo tiene varias hojas. Elige la que corresponda a la jornada que vas a ejecutar; solo se leerá esa.</p>
      <div class="lista">
        ${state.hojas.map((h) => `<div class="item" data-act="elegir-hoja" data-nombre="${esc(h.nombre)}">
          <div class="item-main"><strong>${esc(h.nombre)}</strong>
            <small>${h.filas} filas · seriales: ${esc(h.muestra || '—')}</small></div>
          <span class="chev">›</span></div>`).join('')}
      </div>
    </div>`;
  }

  function viewAreas() {
    if (!state.data) return sinDatos();
    const eds = [...new Set(items().map((i) => i.edificio).filter(Boolean))].sort();
    let lista = items();
    if (state.filtroEdificio) lista = lista.filter((i) => i.edificio === state.filtroEdificio);
    if (state.filtroTexto) {
      const q = clave(state.filtroTexto);
      lista = lista.filter((i) => clave(i.oficina).includes(q) || digitos(i.serial).includes(digitos(state.filtroTexto)));
    }
    const grouped = {};
    for (const it of lista) {
      grouped[it.areaId] ||= { info: it, total: 0, hechos: 0 };
      grouped[it.areaId].total++;
      if (estadoDe(it) !== 'PENDIENTE') grouped[it.areaId].hechos++;
    }
    const arr = Object.values(grouped).sort((a, b) =>
      (a.info.edificio + a.info.piso + a.info.oficina).localeCompare(b.info.edificio + b.info.piso + b.info.oficina));
    return `
      <div class="card tight">
        <select class="filtro" data-act="filtro-ed"><option value="">Todos los edificios</option>
          ${eds.map((e) => `<option ${e === state.filtroEdificio ? 'selected' : ''}>${esc(e)}</option>`).join('')}</select>
        <input class="filtro" data-act="filtro-txt" placeholder="Buscar oficina o serial…" value="${esc(state.filtroTexto)}">
      </div>
      <div class="card"><h2>${arr.length} dependencias · ${lista.length} equipos</h2><div class="lista">
        ${arr.map(({ info, total, hechos }) => `<div class="item" data-act="area" data-id="${esc(info.areaId)}">
          <div class="item-main"><strong>${esc(info.oficina || 'Sin oficina')}</strong>
            <small>${esc(info.piso)} · ${esc(info.edificio)}</small>
            <div class="bar mt"><i class="b-OK" style="width:${(hechos / total) * 100}%"></i></div></div>
          <span class="small muted">${hechos}/${total}</span><span class="chev">›</span></div>`).join('') || '<div class="vacio">Sin resultados</div>'}
      </div></div>`;
  }

  function viewArea(p) {
    const id = decodeURIComponent(p[1] || ''); state.areaActual = id;
    const lista = items().filter((i) => i.areaId === id);
    const s = { TOTAL: lista.length };
    for (const e of ORDEN) s[e] = lista.filter((i) => estadoDe(i) === e).length;
    s.PENDIENTE = s.TOTAL - ORDEN.reduce((a, e) => a + s[e], 0);
    return `
      <div class="card"><h2>${esc(lista[0]?.edificio || '')} · ${esc(lista[0]?.piso || '')}</h2>
        <h3>${esc(lista[0]?.oficina || 'Dependencia')}</h3>
        <div class="mt">${statsHtml(s)}</div>
        <button class="btn primary block mt" data-act="go" data-hash="#/escanear">◎ Escanear en esta área</button></div>
      <div class="card"><h2>Equipos de la dependencia</h2><div class="lista">
        ${lista.map((i) => `<div class="item" data-act="equipo" data-serial="${esc(i.serial)}">
          <div class="item-main"><strong class="mono">${esc(i.serial)}</strong>
            <small>${esc(i.tipo)}${novedadDe(i) ? ' · ' + esc(novedadDe(i)) : ''}</small></div>
          ${chip(estadoDe(i))}<span class="chev">›</span></div>`).join('') || '<div class="vacio">Sin equipos</div>'}
      </div></div>`;
  }

  function viewEscanear() {
    return `<div class="card tight"><div class="small muted">Área actual: <b>${esc(state.areaActual ? areaNombre(state.areaActual) : 'ninguna')}</b></div></div>
      <div id="scannerWrap"><video id="video" playsinline muted autoplay></video><div class="mira"></div></div>
      <div class="card mt"><h2>Escribe el serial si falla la lectura</h2>
        <div class="row"><input class="campo grow mono" id="serialManual" inputmode="numeric" placeholder="Ej. 1020010500072" autocomplete="off">
          <button class="btn primary" data-act="buscar-manual">Buscar</button></div>
        <p class="tiny muted mt">Code 128, EAN-13, QR, Code 39 e ITF. Fecha de ejecución se llena sola; fecha de seguimiento queda en blanco.</p>
      </div>`;
  }

  function viewEquipo(p) {
    const serial = decodeURIComponent(p[1] || '');
    const item = items().find((i) => i.serial === serial);
    if (!item) return `<div class="card"><h3 class="mono">${esc(serial)}</h3>
      <p class="muted">Este serial no aparece en la hoja cargada.</p>
      <div class="btn-grid"><button class="btn primary" data-act="go" data-hash="#/alta/${encodeURIComponent(serial)}">Dar de alta (azul)</button>
      <button class="btn" data-act="go" data-hash="#/escanear">Volver a escanear</button></div></div>`;

    const actual = estadoDe(item);
    state.seleccion = state.seleccionSerial === serial ? state.seleccion : (actual === 'PENDIENTE' ? 'OK' : actual);
    state.seleccionSerial = serial;
    const pr = progreso(state.jornadaId, serial) || {};
    const nota = pr.novedad || item.novedad || '';
    const h = horasDe(item);
    const extraVisible = ['NOVEDAD', 'AUSENTE', 'ALTA'].includes(state.seleccion);
    return `
      <div class="card">
        <div class="row spread"><div class="serial-grande">${esc(serial)}</div>${chip(actual)}</div>
        <div class="small muted mt">${esc(item.tipo || 'Tipo no indicado')}</div>
        <div class="small"><b>En la hoja:</b> ${esc(item.oficina || '')} · ${esc(item.piso || '')} · ${esc(item.edificio || '')}</div>
        <div class="tiny muted mt">Programada: ${esc(item.programada || '—')} · Técnico: ${esc(item.tecnico || '—')} · Fila ${item.filaExcel || '—'}</div>
      </div>
      <div class="card"><h2>¿Qué encontraste?</h2>
        <div class="btn-grid">
          ${['OK', 'NOVEDAD', 'AUSENTE', 'ALTA'].map((e) => `<button class="estado-btn ${e} ${state.seleccion === e ? 'sel' : ''}" data-act="sel" data-estado="${e}">
            ${ESTADOS[e].label}<small>${ESTADOS[e].desc}</small></button>`).join('')}
        </div>
        <div id="camposExtra" class="${extraVisible ? '' : 'hidden'}">
          <label class="campo" for="nota">${state.seleccion === 'NOVEDAD' ? '¿Qué novedad? (obligatorio)' : 'Observaciones (opcional)'}</label>
          <textarea class="campo" id="nota" placeholder="${state.seleccion === 'NOVEDAD' ? 'Ej. pantalla partida…' : 'Ej. se lo llevaron a otra dependencia'}">${esc(nota)}</textarea>
          ${state.seleccion === 'ALTA' ? `<label class="campo" for="destinoArea">Dependencia donde está ahora</label>
            <select class="campo" id="destinoArea">${areas().map((a) => `<option value="${esc(a.id)}" ${a.id === (state.areaActual || item.areaId) ? 'selected' : ''}>${esc(a.oficina || a.piso)} · ${esc(a.piso)} · ${esc(a.edificio)}</option>`).join('')}</select>` : ''}
          ${state.seleccion === 'NOVEDAD' ? `<label class="campo" for="foto">Foto (opcional)</label>
            <input class="campo" type="file" id="foto" accept="image/*" capture="environment"><div id="fotoActual" class="mt"></div>` : ''}
        </div>
        <label class="campo" for="horasTxt">Tiempo ejecutado en este equipo</label>
        <div class="row wrap mb">
          ${[0.5, 1, 1.5, 2].map((v) => `<button class="btn sm ${Number(h) === v ? 'primary' : ''}" data-act="horas" data-valor="${v}">${horasTexto(v)}</button>`).join('')}
        </div>
        <input class="campo" id="horasTxt" value="${esc(pr.horasTexto || horasTexto(h))}" placeholder="Ej. una hora">
        <div class="btn-grid mt">
          <button class="btn primary" data-act="guardar" data-serial="${esc(serial)}" data-siguiente="1">Guardar y seguir</button>
          <button class="btn" data-act="guardar" data-serial="${esc(serial)}">Guardar y salir</button>
        </div>
        <p class="tiny muted mt">Fecha de ejecución: ${hoy()} (automática). Fecha de seguimiento: se deja en blanco.</p>
      </div>`;
  }

  function viewAlta(p) {
    const serial = decodeURIComponent(p[1] || '');
    const eds = [...new Set(areas().map((a) => a.edificio).filter(Boolean))].sort();
    const tipos = [...new Set(state.data?.catalogo.map((e) => e.tipo).filter(Boolean) || [])].sort();
    const actual = areas().find((a) => a.id === state.areaActual);
    const edInicial = actual?.edificio || eds[0];
    return `<div class="card"><h3 class="mono">${esc(serial)}</h3>
      <p class="small muted">Este serial no está en el cronograma. Elige la <b>dependencia donde lo encontraste</b>; se agregará en azul junto a los equipos de esa área.</p>
      <label class="campo" for="altaEd">Edificio</label>
      <select class="campo" id="altaEd" data-act="alta-ed">${eds.map((e) => `<option ${e === edInicial ? 'selected' : ''}>${esc(e)}</option>`).join('')}</select>
      <label class="campo" for="altaArea">Dependencia</label>
      <select class="campo" id="altaArea">${areas().filter((a) => a.edificio === edInicial)
        .map((a) => `<option value="${esc(a.id)}" ${a.id === state.areaActual ? 'selected' : ''}>${esc(a.oficina || a.piso)} · ${esc(a.piso)}</option>`).join('')}</select>
      <label class="campo" for="altaTipo">Tipo de equipo</label>
      <select class="campo" id="altaTipo">${tipos.map((t) => `<option>${esc(t)}</option>`).join('')}<option>OTRO</option></select>
      <label class="campo" for="altaNota">Observaciones (opcional)</label>
      <input class="campo" id="altaNota" placeholder="Ej. equipo nuevo / reasignado desde Rectoría">
      <label class="campo" for="horasTxt">Tiempo ejecutado</label>
      <input class="campo" id="horasTxt" value="una hora">
      <button class="btn primary block mt" data-act="confirmar-alta" data-serial="${esc(serial)}">Agregar en azul</button></div>`;
  }

  function viewDatos() {
    const a = state.data?.alertas || []; const s = stats(); const j = jornada(); const sy = leer(K.sync, null);
    return `
      <div class="card"><h2>Informe final</h2>
        <p class="small muted">Genera un Excel con el mismo formato del archivo que subiste: fecha de ejecución de hoy, seguimiento en blanco, horas, observaciones y colores. ${s.TOTAL - s.HECHOS} equipos aún pendientes.</p>
        <button class="btn primary block" data-act="informe">↧ Generar informe final (.xlsx)</button>
        <button class="btn block mt" data-act="exportar">Exportar planilla de control (auditoría)</button>
        <button class="btn block mt" data-act="subir">⇧ Subir / cambiar cronograma</button>
        <div class="small muted mt">Archivo: ${esc(state.archivoMeta?.nombre || state.data?.fuente || '—')}<br>Hoja: ${esc(state.archivoMeta?.hoja || '—')}</div>
      </div>
      <div class="card"><h2>Google Sheets (opcional)</h2>
        <p class="small muted">Cada equipo se envía <b>al dar Guardar y seguir</b> a la hoja <b>${esc(nombreHoja())}</b>. Además, al abrir la app y cada 2 min se envía lo pendiente y se <b>traen los cambios</b> que hicieron en el Sheet (gana lo más reciente). Al escanear también se actualiza. Sin internet queda en cola.</p>
        ${pendientes().length ? `<p class="small" style="color:#b45309">⏳ ${pendientes().length} registro(s) esperando para enviar</p>` : '<p class="small muted">✓ Sin registros pendientes</p>'}
        <p class="small muted">↻ Auto-sync: al abrir, al escanear, al volver la señal y cada 2 min. Última: ${esc(sy?.hora || '—')}${sy?.cambios ? ' · ' + sy.cambios + ' cambio(s) traídos' : ''}</p>
        ${sy?.coloreadas !== undefined ? `<p class="small muted">Sheet leído: ${sy.coloreadas} con color · ${sy.conDatos} con datos (fecha/horas/observaciones)</p>` : ''}
        <label class="campo" for="cfgUrl">URL del Web App</label>
        <input class="campo" id="cfgUrl" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(state.config.url || '')}">
        <label class="campo" for="cfgToken">Clave secreta</label>
        <input class="campo" id="cfgToken" value="${esc(state.config.token || '')}">
        <label class="campo" for="cfgTecnico">Técnico</label>
        <input class="campo" id="cfgTecnico" placeholder="TU NOMBRE" value="${esc(state.config.tecnico || '')}">
        <div class="btn-grid mt"><button class="btn" data-act="guardar-config">Guardar datos</button>
          <button class="btn primary" data-act="sincronizar">Sincronizar ahora</button></div>
        <div id="syncResult" class="small mt"></div>
      </div>
      ${a.length ? `<div class="card"><h2>Incidencias del Excel (${a.length})</h2>
        ${a.slice(0, 30).map((x) => `<div class="alerta">Fila ${esc(x.fila)} · <span class="mono">${esc(x.serial)}</span> · ${esc(x.motivo)}</div>`).join('')}</div>` : ''}
      <div class="card"><h2>Zona de cuidado</h2>
        <button class="btn block" data-act="borrar-progreso">Borrar registros de esta jornada</button>
        <button class="btn block mt" data-act="olvidar">Olvidar cronograma y volver al inicial</button></div>`;
  }

  function viewAyuda() {
    return `<div class="card"><h2>Flujo de trabajo</h2><ol class="small">
      <li>Sube el Excel del mantenimiento. Si tiene varias hojas, elige la que te corresponde.</li>
      <li>Entra a Áreas y elige la dependencia por donde vas a empezar.</li>
      <li>Escanea el código de barras. La app muestra edificio, piso, oficina y tipo.</li>
      <li>Toca el estado: hecho, novedad, no se encuentra u otra dependencia/nuevo.</li>
      <li>Si es novedad, escribe qué pasó y opcionalmente toma foto. Si cambió de dependencia, confirma dónde está.</li>
      <li>Ajusta el tiempo ejecutado y toca <b>Guardar y seguir</b>.</li>
      <li>La fecha de ejecución se pone sola; la de seguimiento queda vacía para la otra persona.</li>
      <li>Al terminar, genera el <b>informe final</b> y entrégalo.</li>
    </ol><hr><h2>Colores</h2><div class="lista">
      ${['OK', 'NOVEDAD', 'AUSENTE', 'ALTA'].map((e) => `<div class="item"><div class="item-main"><strong>${ESTADOS[e].label}</strong><small>${ESTADOS[e].desc}</small></div>${chip(e)}</div>`).join('')}
    </div><p class="tiny muted mt">Todo se guarda en el teléfono. Sincroniza cuando tengas internet.</p></div>`;
  }

  function viewEtiquetas() {
    const eds = [...new Set(items().map((i) => i.edificio).filter(Boolean))].sort();
    return `<div class="card"><h2>Generar etiquetas Code 128</h2>
      <p class="small muted">Para equipos que no tengan código legible. Filtra, genera e imprime en papel adhesivo.</p>
      <select class="filtro" id="etiEd"><option value="">Todos los edificios</option>${eds.map((e) => `<option>${esc(e)}</option>`).join('')}</select>
      <button class="btn primary block" data-act="generar-etiquetas">Generar etiquetas</button><div id="etiquetas" class="mt"></div></div>`;
  }

  /* ---------------- acciones de registro ---------------- */

  /** Inserta un equipo justo después del último de su área (no al final). */
  function insertarItemEnArea(item, areaId) {
    const j = jornada(); if (!j) return;
    const viejo = j.items.findIndex((i) => i.serial === item.serial);
    if (viejo >= 0) j.items.splice(viejo, 1);
    const norm = (v) => String(v || '').trim().toUpperCase()
      .replace(/[ÁÀÄÅ]/g, 'A').replace(/[ÉÈË]/g, 'E').replace(/[ÍÌÏ]/g, 'I')
      .replace(/[ÓÒÖ]/g, 'O').replace(/[ÚÙÜ]/g, 'U').replace(/Ñ/g, 'N');
    const esMisma = (i) => i.areaId === areaId || (
      norm(i.oficina) && norm(i.oficina) === norm(item.oficina)
      && norm(i.edificio) === norm(item.edificio)
      && norm(i.piso) === norm(item.piso)
    );
    let ultima = -1;
    j.items.forEach((i, n) => { if (esMisma(i)) ultima = n; });
    j.items.splice(ultima >= 0 ? ultima + 1 : j.items.length, 0, item);
  }

  function setEstado(serial, datos) {
    if (!state.progress[state.jornadaId]) state.progress[state.jornadaId] = {};
    state.progress[state.jornadaId][serial] = {
      estado: datos.estado,
      novedad: datos.novedad || '',
      horas: Number(datos.horas) || 1,
      horasTexto: datos.horasTexto || horasTexto(datos.horas),
      fecha: hoy() + ' ' + hora().slice(11),
      tecnico: datos.tecnico || state.config.tecnico || jornada()?.tecnico || '',
      registrado: true, // solo lo que el usuario grabó con Guardar y seguir
      ...(datos.extra || {}),
    };
    guardarProgreso();
  }

  function leerHoras() {
    const txt = norm($('#horasTxt')?.value || 'una hora');
    return { horas: horasDeTexto(txt) || 1, horasTexto: txt || 'una hora' };
  }

  function observacionFinal(estado, nota, item) {
    if (estado === 'OK') return 'sin novedades';
    if (estado === 'NOVEDAD') return nota;
    if (estado === 'AUSENTE') return /no se encuentra/i.test(nota) ? nota : ('no se encuentra' + (nota ? '. ' + nota : ''));
    if (estado === 'ALTA') return nota || `Ubicado en ${areaNombre(item.areaId)}`;
    return nota;
  }

  async function guardar(serial, seguir) {
    const item = items().find((i) => i.serial === serial);
    if (!item) { toast('Equipo no encontrado'); return; }
    const estado = (state.seleccion && state.seleccion !== 'PENDIENTE') ? state.seleccion : estadoDe(item);
    let nota = $('#nota')?.value || '';
    if (estado === 'NOVEDAD' && !norm(nota)) { toast('Escribe cuál es la novedad'); return; }
    const { horas, horasTexto: ht } = leerHoras();
    const extra = {};

    if (estado === 'ALTA') {
      const destino = $('#destinoArea')?.value || state.areaActual || item.areaId;
      const area = areas().find((a) => a.id === destino);
      if (area) {
        if (item.areaId !== destino) {
          extra.origen = areaNombre(item.areaId);
          item.areaId = destino; item.edificio = area.edificio; item.piso = area.piso; item.oficina = area.oficina;
          insertarItemEnArea(item, destino); // se mueve a su área nueva, no al final
          const cat = state.data.catalogo.find((c) => c.serial === serial);
          if (cat) Object.assign(cat, { areaId: destino, edificio: area.edificio, piso: area.piso, oficina: area.oficina });
        }
        nota = norm(nota) || (extra.origen ? `Reasignado desde ${extra.origen}` : `Ubicado en ${areaNombre(destino)}`);
        guardarDataset();
      }
    }

    const foto = $('#foto')?.files?.[0];
    if (foto) {
      const dataUrl = await new Promise((res) => {
        const fr = new FileReader(); fr.onload = () => res(fr.result); fr.readAsDataURL(foto);
      });
      await fotoGuardar(serial, dataUrl);
    }
    setEstado(serial, { estado, novedad: observacionFinal(estado, nota, item), horas, horasTexto: ht, extra });
    vibrar(120);
    state.seleccion = null; state.seleccionSerial = null;
    if (norm(state.config.url)) {
      const enviado = await enviarFila(serial);
      if (enviado) leerDesdeSheets(state.config.url, state.config.token);
      toast(`${ESTADOS[estado].label} · ${serial} · ${enviado ? 'enviado a Sheets' : 'sin internet: se enviará después'}`);
    } else {
      toast(`${ESTADOS[estado].label} · ${serial}`);
    }
    go(seguir ? '#/escanear' : '#/area/' + encodeURIComponent(item.areaId));
  }

  async function darDeAlta(serial) {
    const areaId = $('#altaArea')?.value;
    const area = areas().find((a) => a.id === areaId);
    if (!area) { toast('Elige una dependencia'); return; }
    const tipo = $('#altaTipo')?.value || 'OTRO';
    const { horas, horasTexto: ht } = leerHoras();
    const nota = $('#altaNota')?.value || `Equipo nuevo / ubicado en ${areaNombre(areaId)}`;
    const j = jornada();
    if (!state.data.catalogo.some((e) => e.serial === serial)) {
      state.data.catalogo.push({ serial, areaId, edificio: area.edificio, piso: area.piso, oficina: area.oficina, tipo, serialValido: /^\d+$/.test(serial) });
    }
    let item = j.items.find((i) => i.serial === serial);
    if (!item) {
      item = { serial, areaId, edificio: area.edificio, piso: area.piso, oficina: area.oficina, tipo,
        estado: 'ALTA', novedad: nota, programada: '', tecnico: state.config.tecnico || j.tecnico || '',
        duracion: ht, filaExcel: null, serialValido: /^\d+$/.test(serial) };
    } else {
      item.areaId = areaId; item.edificio = area.edificio; item.piso = area.piso; item.oficina = area.oficina;
      item.tipo = tipo;
    }
    insertarItemEnArea(item, areaId); // va donde está su área, no al final
    state.areaActual = areaId;
    setEstado(serial, { estado: 'ALTA', novedad: nota, horas, horasTexto: ht });
    guardarDataset(); vibrar(120);
    if (norm(state.config.url)) {
      const enviado = await enviarFila(serial);
      if (enviado) leerDesdeSheets(state.config.url, state.config.token);
      toast(enviado ? 'Agregado en azul · enviado a Sheets' : 'Agregado en azul · sin internet: se enviará después');
    } else {
      toast('Agregado en azul');
    }
    go('#/equipo/' + encodeURIComponent(serial));
  }

  /** Cada escaneo/búsqueda reenvía pendientes y trae lo último del Sheet. */
  function actualizarPorScan() {
    if (!norm(state.config.url)) return;
    reintentarPendientes().then(() => leerDesdeSheets(state.config.url, state.config.token));
  }

  function buscarSerial(texto) {
    const t = norm(texto); if (!t) return;
    actualizarPorScan();
    const dt = digitos(t);
    const item = items().find((i) => i.serial === t)
      || (dt.length >= 8 && items().find((i) => digitos(i.serial) === dt))
      || (dt.length >= 8 && items().find((i) => digitos(i.serial).slice(-dt.length) === dt));
    if (item) {
      // El área actual es la dependencia que se está recorriendo, no la del equipo.
      state.seleccion = null; state.seleccionSerial = null;
      go('#/equipo/' + encodeURIComponent(item.serial));
    } else {
      vibrar([60, 60, 60]); toast('Serial no está en la hoja cargada');
      go('#/alta/' + encodeURIComponent(t));
    }
  }

  /* ---------------- escáner ---------------- */

  let codeReader = null, nativeLoop = false;
  function detenerEscanner() {
    try { codeReader?.reset(); } catch (_) {}
    codeReader = null; nativeLoop = false;
    const v = $('#video');
    if (v?.srcObject) { v.srcObject.getTracks().forEach((t) => t.stop()); v.srcObject = null; }
  }
  function iniciarEscanner() {
    const video = $('#video'); if (!video) return;
    if (!navigator.mediaDevices?.getUserMedia) { toast('Este navegador no permite cámara. Usa el campo manual.'); return; }
    if (Date.now() - state.ultimoScan.hora < 2500) return;
    state.ultimoScan = { texto: '', hora: Date.now() };
    const onResult = (texto) => {
      const t = norm(texto); if (!t || t === state.ultimoScan.texto) return;
      if (Date.now() - state.ultimoScan.hora < 2500) return;
      state.ultimoScan = { texto: t, hora: Date.now() }; vibrar(80); buscarSerial(t);
    };
    if (window.ZXing?.BrowserMultiFormatReader) {
      try {
        const hints = new Map([[ZXing.DecodeHintType.POSSIBLE_FORMATS, [
          ZXing.BarcodeFormat.CODE_128, ZXing.BarcodeFormat.CODE_39, ZXing.BarcodeFormat.ITF,
          ZXing.BarcodeFormat.EAN_13, ZXing.BarcodeFormat.QR_CODE,
        ]]]);
        codeReader = new ZXing.BrowserMultiFormatReader(hints);
        const leer = (r) => { if (r) onResult(r.getText()); };
        codeReader.decodeFromConstraints({ video: { facingMode: { ideal: 'environment' } } }, video, leer)
          .catch(() => codeReader.decodeFromVideoDevice(null, video, leer))
          .catch(() => iniciarNativo(video, onResult));
        return;
      } catch (_) {}
    }
    iniciarNativo(video, onResult);
  }
  async function iniciarNativo(video, onResult) {
    if (nativeLoop) return;
    if (!('BarcodeDetector' in window)) { toast('No se pudo abrir el lector. Usa el campo manual.'); return; }
    const detector = new BarcodeDetector({ formats: ['code_128', 'code_39', 'ean_13', 'qr_code', 'itf'] });
    nativeLoop = true;
    const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
    video.srcObject = stream; await video.play();
    const tick = async () => {
      if (!nativeLoop) return;
      try { const c = await detector.detect(video); if (c[0]) onResult(c[0].rawValue); } catch (_) {}
      requestAnimationFrame(tick);
    };
    tick();
  }

  /* ---------------- importar Excel con elección de hoja ---------------- */

  async function importar(tipo) {
    const inp = $('#archivo') || document.createElement('input');
    if (!inp.id) { inp.id = 'archivo'; inp.type = 'file'; inp.className = 'hidden'; document.body.appendChild(inp); }
    inp.accept = tipo === 'xlsx' ? '.xlsx,.xlsm' : '.json';
    inp.value = '';
    inp.onchange = async () => {
      const f = inp.files?.[0]; if (!f) return;
      try {
        if (tipo === 'json') {
          const data = JSON.parse(await f.text());
          if (!data.jornadas?.length) throw new Error('JSON sin jornadas');
          state.data = data; state.archivoMeta = null;
          localStorage.removeItem(K.archivo); guardarDataset();
          go('#/inicio'); location.reload(); return;
        }
        const buffer = await f.arrayBuffer();
        await archivoGuardar(new Blob([buffer]));
        const wb = XLSX.read(buffer, { type: 'array', cellStyles: true });
        const candidatas = wb.SheetNames.filter((n) => {
          const filas = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, blankrows: false }).slice(0, 12);
          return filas.some((r) => r.some((v) => /serial/i.test(String(v ?? ''))));
        });
        const nombres = candidatas.length ? candidatas : wb.SheetNames;
        state.hojas = nombres.map((n) => {
          const ws = wb.Sheets[n];
          const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1');
          const filas = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false }).slice(0, 40);
          const muestra = filas.map((r) => r.find((v) => /^\d{8,}$/.test(String(v ?? '')))).filter(Boolean).slice(0, 2).join(', ');
          return { nombre: n, filas: ref.e.r + 1, columnas: ref.e.c + 1, muestra };
        });
        state.archivoMeta = { nombre: f.name, hoja: '' };
        if (state.hojas.length === 1) await usarHoja(state.hojas[0].nombre);
        else go('#/elegir-hoja');
      } catch (e) { toast('No se pudo leer el archivo: ' + e.message); }
    };
    inp.click();
  }

  async function usarHoja(nombre) {
    try {
      const blob = await archivoLeer();
      if (!blob) throw new Error('Falta el archivo');
      state.data = parseXlsx(await blob.arrayBuffer(), nombre);
      state.archivoMeta = { nombre: state.archivoMeta?.nombre || 'cronograma.xlsx', hoja: nombre };
      guardarDataset(); escribir(K.archivo, state.archivoMeta);
      state.jornadaId = state.data.jornadas[0].id;
      escribir(K.jornada, state.jornadaId);
      state.progress[state.jornadaId] = {}; // cronograma nuevo: parte sin registros
      guardarProgreso();
      toast(`Hoja "${nombre}": ${state.data.resumen.equiposUnicos} equipos. Subiendo a Sheets...`);
      await subirCronograma(nombre);
      go('#/inicio'); location.reload();
    } catch (e) { toast('No se pudo leer la hoja: ' + e.message); }
  }

  function detectarColumnas(ws, ref, val) {
    const map = { ...COLUMNAS_POR_DEFECTO };
    let fila = -1;
    for (let r = ref.s.r; r <= Math.min(ref.s.r + 15, ref.e.r); r++) {
      for (let c = ref.s.c; c <= ref.e.c; c++) {
        if (/serial/i.test(val(r, c))) { fila = r; break; }
      }
      if (fila >= 0) break;
    }
    if (fila < 0) return map;
    const patrones = {
      edificio: /edificio|sede|bloque/i, piso: /piso|nivel/i, oficina: /oficina|dependencia|area/i,
      serial: /serial|inventario/i, tipo: /tipo.*equipo|equipo/i, programada: /fecha.*programada|programad/i,
      ejecutada: /fecha.*ejecutada|ejecutad/i, tecnico: /tecnico|técnico/i, semestre: /semestre/i,
      duracion: /tiempo|duracion|duración/i, seguimiento: /seguimiento/i, observaciones: /observacion|observación|novedad/i,
    };
    map.headerRow = fila;
    for (let c = ref.s.c; c <= ref.e.c; c++) {
      const h = val(fila, c);
      if (!h) continue;
      for (const [k, re] of Object.entries(patrones)) {
        if (re.test(h) && map[k] === COLUMNAS_POR_DEFECTO[k]) map[k] = c;
      }
    }
    return map;
  }

  function parseXlsx(buffer, nombreHoja) {
    const wb = XLSX.read(buffer, { type: 'array', cellStyles: true, cellDates: true });
    const ws = wb.Sheets[nombreHoja];
    if (!ws) throw new Error('No existe la hoja ' + nombreHoja);
    const merged = {};
    for (const m of (ws['!merges'] || [])) {
      const v = ws[XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })]?.v;
      if (v == null) continue;
      for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) merged[r + ':' + c] = v;
    }
    const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const val = (r, c) => {
      const v = ws[XLSX.utils.encode_cell({ r, c })]?.v ?? merged[r + ':' + c];
      return v instanceof Date ? v.toISOString().slice(0, 10) : norm(v);
    };
    const col = detectarColumnas(ws, ref, val);
    const cat = {}, areas = {}, items = [], alertas = [], vistos = {};
    for (let r = col.headerRow + 1; r <= ref.e.r; r++) {
      const edificio = val(r, col.edificio), piso = val(r, col.piso), oficina = val(r, col.oficina);
      const serial = val(r, col.serial), tipo = val(r, col.tipo);
      if (!serial && !tipo && !oficina) continue;
      const valido = /^\d{9,16}$/.test(serial);
      if (!valido) {
        // Sin serial no hay equipo: se reporta en Alertas pero NO cuenta en las estadísticas.
        alertas.push({ jornada: nombreHoja, fila: r + 1, serial, motivo: 'Serial no numérico o vacío' });
        continue;
      }
      if (vistos[serial] !== undefined) alertas.push({ jornada: nombreHoja, fila: r + 1, serial, motivo: 'Serial duplicado (fila ' + (vistos[serial] + 1) + ')' });
      else vistos[serial] = r;
      const areaId = clave([edificio, piso, oficina].join('|')) || 'sin-area';
      areas[areaId] ||= { id: areaId, edificio, piso, oficina };
      const s = serial;
      const obs = val(r, col.observaciones);
      // SheetJS guarda el relleno en s.fgColor (algunas versiones, en s.fill.fgColor).
      // Sin color reconocido el equipo sigue PENDIENTE, aunque tenga fecha o semestre.
      const estilo = ws[XLSX.utils.encode_cell({ r, c: col.serial })]?.s || {};
      let crudo = String(estilo.fgColor?.rgb || estilo.fill?.fgColor?.rgb || '')
        .toUpperCase().replace(/^0X/, '');
      if (crudo.length === 8) crudo = crudo.slice(-6);
      const color = crudo;
      let estado = 'PENDIENTE';
      if (color === '00B0F0') estado = 'ALTA';
      else if (color === 'FF0000' || color === 'FF3300' || /no se encuentra/i.test(obs)) estado = 'AUSENTE';
      else if (color === 'FFC000' || color === 'ED7D31' || (obs && !/^sin novedad/i.test(obs))) estado = 'NOVEDAD';
      else if (color === '92D050' || color === '70AD47') estado = 'OK';
      if (!cat[s]) cat[s] = { serial: s, areaId, edificio, piso, oficina, tipo, serialValido: valido };
      items.push({ serial: s, areaId, edificio, piso, oficina, tipo, estado,
        novedad: estado === 'NOVEDAD' ? obs : '', programada: val(r, col.programada).slice(0, 10),
        ejecutada: val(r, col.ejecutada).slice(0, 10), tecnico: val(r, col.tecnico),
        semestre: val(r, col.semestre), duracion: val(r, col.duracion), filaExcel: r + 1, serialValido: valido });
    }
    if (!items.length) throw new Error('La hoja no tiene equipos');
    const anio = (nombreHoja.match(/\d{4}/) || [new Date().getFullYear()])[0];
    const sem = (nombreHoja.match(/[AB]$/i) || [''])[0].toUpperCase();
    return {
      generado: hora(), fuente: state.archivoMeta?.nombre || nombreHoja, columnas: col,
      areas: Object.values(areas), catalogo: Object.values(cat),
      jornadas: [{ id: clave(nombreHoja), nombre: `Preventivo ${sem} ${anio}`.trim(), semestre: sem, anio,
        tecnico: items.find((i) => i.tecnico)?.tecnico || '',
        items }],
      alertas,
      resumen: { equiposUnicos: Object.keys(cat).length, areas: Object.keys(areas).length, jornadas: 1, registros: items.length, alertas: alertas.length },
    };
  }

  /* ---------------- informe final ---------------- */

  async function generarInforme() {
    if (!state.data) { toast('No hay cronograma'); return; }
    const btn = document.querySelector('[data-act=informe]');
    if (btn) { btn.disabled = true; btn.textContent = 'Generando…'; }
    try {
      const j = jornada(); const col = columnas(); const s = stats();
      let wb;
      const blob = await archivoLeer();
      if (blob && window.ExcelJS) {
        wb = new ExcelJS.Workbook(); await wb.xlsx.load(await blob.arrayBuffer());
      } else {
        wb = new ExcelJS.Workbook();
      }
      let ws = wb.getWorksheet(nombreHoja());
      if (!ws) ws = wb.addWorksheet(nombreHoja());

      // El cronograma original tiene columnas completas unidas (A2:A265, G2:G21, K2:K21…).
      // Para poder escribir fecha/observaciones por equipo hay que separar esas celdas.
      const escribir = [col.edificio, col.piso, col.oficina, col.serial, col.tipo, col.ejecutada,
        col.tecnico, col.semestre, col.duracion, col.observaciones];
      for (const ref of [...(ws.model.merges || [])]) {
        try {
          const r = XLSX.utils.decode_range(ref);
          if (escribir.some((c) => c >= r.s.c && c <= r.e.c)) ws.unMergeCells(ref);
        } catch (_) {}
      }

      const encabezados = ['EDIFICIO', 'PISO', 'OFICINA', 'SERIAL UDES', 'TIPO EQUIPO', 'FECHA PROGRAMADA',
        'FECHA EJECUTADA', 'TÉCNICO', 'SEMESTRE', 'TIEMPO EJECUCIÓN', 'FECHA SEGUIMIENTO', 'OBSERVACIONES'];
      encabezados.forEach((t, i) => {
        if (!ws.getRow(col.headerRow + 1).getCell(i + 1).value) ws.getRow(col.headerRow + 1).getCell(i + 1).value = t;
      });

      let registrados = 0;
      let filaNueva = ws.rowCount + 1; // equipos de alta no tenían fila en el Excel original
      for (const it of items()) {
        const pr = progreso(j.id, it.serial) || {};
        const est = pr.registrado ? pr.estado : 'PENDIENTE';
        if (est === 'PENDIENTE') continue;
        const fila = ws.getRow(it.filaExcel || filaNueva++);
        fila.getCell(col.edificio + 1).value = it.edificio;
        fila.getCell(col.piso + 1).value = it.piso;
        fila.getCell(col.oficina + 1).value = it.oficina;
        fila.getCell(col.serial + 1).value = it.serial;
        fila.getCell(col.tipo + 1).value = it.tipo;
        fila.getCell(col.ejecutada + 1).value = hoy();
        fila.getCell(col.tecnico + 1).value = pr.tecnico || state.config.tecnico || it.tecnico || j.tecnico || '';
        fila.getCell(col.semestre + 1).value = pr.semestre || it.semestre || j.semestre || '';
        fila.getCell(col.duracion + 1).value = pr.horasTexto || horasTexto(pr.horas || 1);
        fila.getCell(col.seguimiento + 1).value = null;
        fila.getCell(col.observaciones + 1).value = pr.novedad || observacionFinal(est, '', it);
        const c = fila.getCell(col.serial + 1);
        // ExcelJS comparte el objeto de estilo entre celdas cargadas del XLSX;
        // sin esta copia, el último color sobrescribe todos los anteriores.
        c.style = { ...c.style };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ESTADOS[est].argb } };
        // Solo cambia el fondo; el número queda en negro y se respeta el resto de la fuente.
        c.font = { ...(c.font || {}), color: { argb: 'FF000000' } };
        registrados++;
      }

      if (wb.getWorksheet('RESUMEN')) wb.removeWorksheet('RESUMEN');
      const res = wb.addWorksheet('RESUMEN');
      res.addRows([
        ['INFORME DE MANTENIMIENTO PREVENTIVO'],
        ['Hoja', nombreHoja()],
        ['Archivo', state.archivoMeta?.nombre || state.data.fuente || ''],
        ['Técnico', state.config.tecnico || j.tecnico || ''],
        ['Generado', hora()],
        [],
        ['ESTADO', 'CANTIDAD', 'COLOR'],
        ...['OK', 'NOVEDAD', 'AUSENTE', 'ALTA', 'PENDIENTE'].map((e) => [ESTADOS[e].label, s[e] || 0, '']),
        [],
        ['TOTAL EQUIPOS', s.TOTAL],
        ['REGISTRADOS', registrados],
        ['PENDIENTES', s.PENDIENTE || 0],
        ['TOTAL HORAS', Number(s.HORAS.toFixed(1))],
        ['FECHA EJECUCIÓN', 'Automática: ' + hoy()],
        ['FECHA SEGUIMIENTO', 'En blanco (la llena otra persona)'],
      ]);
      res.getColumn(1).width = 30; res.getColumn(2).width = 42; res.getColumn(3).width = 14;
      res.getCell(1, 1).font = { bold: true, size: 14 };
      res.getCell(7, 1).font = { bold: true };

      const buf = await wb.xlsx.writeBuffer();
      descargar(buf, `informe_${j.id}_${hoy()}.xlsx`,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      toast(`Informe listo · ${registrados} equipos`);
    } catch (e) {
      toast('No se pudo generar: ' + e.message);
      console.error(e);
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = '↧ Generar informe final (.xlsx)'; }
    }
  }

  function exportar() {
    const j = jornada(); if (!j) return;
    const filas = items().map((i) => {
      const pr = progreso(j.id, i.serial) || {};
      const est = pr.estado || i.estado || 'PENDIENTE';
      return { Jornada: j.nombre, Serial: i.serial, Edificio: i.edificio, Piso: i.piso, Oficina: i.oficina,
        Tipo: i.tipo, Estado: est, EstadoTexto: ESTADOS[est]?.label || est,
        Novedad: pr.novedad || i.novedad || '', Horas: pr.horas || 1, HorasTexto: pr.horasTexto || '',
        Tecnico: pr.tecnico || state.config.tecnico || i.tecnico || '', FechaEjecucion: est === 'PENDIENTE' ? '' : hoy(),
        FechaSeguimiento: '', FilaExcel: i.filaExcel || '' };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas), 'Ejecuciones');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      filas.filter((f) => ['ALTA', 'AUSENTE'].includes(f.Estado))), 'Cambios');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      (state.data.alertas || []).map((a) => ({ Jornada: a.jornada, Fila: a.fila, Serial: a.serial, Motivo: a.motivo }))), 'Alertas');
    XLSX.writeFile(wb, `control_${j.id}_${hoy()}.xlsx`); toast('Planilla de control descargada');
  }

  function filaPayload(i, pr, j, soloRegistrados) {
    // En los envíos al Sheet solo cuenta lo que se grabó en el teléfono.
    // El color verde del Excel original NO significa que el equipo fue ejecutado.
    const est = soloRegistrados ? (pr?.registrado ? pr.estado : 'PENDIENTE')
      : (pr?.estado || i.estado || 'PENDIENTE');
    return { serial: i.serial, estado: est, edificio: i.edificio, piso: i.piso, oficina: i.oficina,
      tipo: i.tipo, programada: i.programada || '', semestre: i.semestre || j.semestre || '',
      fecha: est === 'PENDIENTE' ? '' : (pr?.fecha || hoy()),
      tecnico: pr?.tecnico || state.config.tecnico || j.tecnico || '',
      horas: pr?.horas || 1, horasTexto: pr?.horasTexto || '', seguimiento: '',
      observaciones: pr?.novedad || observacionFinal(est, '', i), color: ESTADOS[est]?.argb || ESTADOS.PENDIENTE.argb };
  }

  /** POST al Web App con límite de tiempo: si la señal se cuelga, se encola el registro. */
  function postSheet(url, body) {
    return new Promise((res, rej) => {
      const t = setTimeout(() => rej(new Error('Tiempo de espera agotado')), 8000);
      fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body })
        .then(() => { clearTimeout(t); res(); })
        .catch((e) => { clearTimeout(t); rej(e); });
    });
  }

  /** Filas con el estado actual, para aplicarlas encima del cronograma recién subido. */
  function filasActuales() {
    const j = jornada(); if (!j) return [];
    return items().map((i) => filaPayload(i, progreso(j.id, i.serial), j, true));
  }

  /** Convierte la hoja original del Excel en un payload que Apps Script puede recrear. */
  async function serializarHoja(nombreHoja) {
    const blob = await archivoLeer();
    if (!blob) return null;
    const wb = XLSX.read(await blob.arrayBuffer(), { type: 'array', cellStyles: true, cellDates: true });
    const ws = wb.Sheets[nombreHoja];
    if (!ws) return null;
    const ref = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const nFil = ref.e.r + 1, nCol = ref.e.c + 1;
    const valores = [], tipos = [], estilos = [];
    for (let r = 0; r < nFil; r++) {
      const vf = [], tf = [], ef = [];
      for (let c = 0; c < nCol; c++) {
        const cel = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cel) { vf.push(''); tf.push(''); ef.push({}); continue; }
        let v = cel.v;
        if (v instanceof Date) v = v.toISOString().slice(0, 10);
        vf.push(v == null ? '' : v);
        tf.push(cel.t === 'd' ? 'd' : '');
        const s = cel.s || {};
        const bg = s.patternType === 'solid' && s.fgColor ? String(s.fgColor.rgb || '') : '';
        const fg = s.font?.color?.rgb ? String(s.font.color.rgb) : '';
        ef.push({ bg: bg.replace(/^0x/, ''), fg: fg.replace(/^0x/, '') });
      }
      valores.push(vf); tipos.push(tf); estilos.push(ef);
    }
    const merges = (ws['!merges'] || []).map((m) => [m.s.r, m.s.c, m.e.r, m.e.c]);
    const cols = (ws['!cols'] || []).slice(0, nCol).map((c) => c ? Math.round(c.wpx || ((c.width || 10) * 7) + 5) : 0);
    const rows = (ws['!rows'] || []).slice(0, nFil).map((h) => h ? Math.round(h.hpx || ((h.hpt || 15) * 1.33)) : 0);
    while (cols.length < nCol) cols.push(0);
    while (rows.length < nFil) rows.push(0);
    const col = state.data?.columnas || COLUMNAS_POR_DEFECTO;
    return { valores, tipos, estilos, merges, cols, rows, congelar: (col.headerRow || 0) + 1 };
  }

  /** Reemplaza el Google Sheet completo por la hoja del cronograma que se acaba de cargar. */
  async function subirCronograma(nombre) {
    const url = norm(state.config.url);
    if (!url || !nombre) { escribir(K.cargar, { hoja: nombre }); return false; }
    try {
      const base = await serializarHoja(nombre);
      if (!base) throw new Error('No se pudo leer el Excel original');
      const j = jornada();
      await postSheet(url, JSON.stringify({
        token: state.config.token || 'udes2026',
        hoja: nombre, jornada: j ? j.id : nombre, fecha: hora(),
        tecnico: state.config.tecnico || j?.tecnico || '',
        columnas: state.data?.columnas || COLUMNAS_POR_DEFECTO,
        total: items().length, modo: 'cargar',
        ...base, filas: filasActuales(),
      }));
      localStorage.removeItem(K.cargar);
      return true;
    } catch (e) {
      escribir(K.cargar, { hoja: nombre, error: e.message });
      return false;
    }
  }

  function pendientes() { return leer(K.pendiente, []); }
  function marcarPendiente(serial) {
    const p = pendientes();
    if (p.indexOf(serial) === -1) { p.push(serial); escribir(K.pendiente, p); }
  }
  function limpiarPendiente(serial) { escribir(K.pendiente, pendientes().filter((s) => s !== serial)); }

  /** Envía UN solo equipo al Google Sheet en cuanto se guarda. Devuelve true si llegó. */
  async function enviarFila(serial) {
    const url = norm(state.config.url);
    if (!url) return false;
    const j = jornada();
    const item = items().find((i) => i.serial === serial);
    if (!item) return false;
    try {
      await postSheet(url, JSON.stringify({ token: state.config.token,
        hoja: nombreHoja(), jornada: j.id, fecha: hora(),
        tecnico: state.config.tecnico || j.tecnico || '', columnas: columnas(),
        total: items().length, modo: 'fila', filas: [filaPayload(item, progreso(j.id, serial), j, true)] }));
      limpiarPendiente(serial);
      return true;
    } catch (e) { marcarPendiente(serial); return false; }
  }

  async function reintentarPendientes() {
    if (!pendientes().length || !norm(state.config.url)) return;
    for (const s of pendientes().slice()) await enviarFila(s);
  }

  /** Lee el Google Sheet mediante JSONP (Apps Script no manda cabeceras CORS). */
  function leerDesdeSheets(url, token) {
    const destino = norm(url || state.config.url);
    if (!destino || !state.data) return;
    const cb = '__prevPull';
    window[cb] = aplicarCambiosSheet;
    const sep = destino.indexOf('?') >= 0 ? '&' : '?';
    const src = destino + sep + 'token=' + encodeURIComponent(token || state.config.token || 'udes2026')
      + '&hoja=' + encodeURIComponent(nombreHoja())
      + '&callback=' + cb + '&t=' + Date.now();
    const s = document.createElement('script');
    s.src = src;
    s.onerror = () => { escribir(K.sync, { hora: hora(), ok: false, error: 'No se pudo leer el Sheet' }); };
    document.head.appendChild(s);
    setTimeout(() => s.remove(), 15000);
  }

  /** Aplica al teléfono lo que cambiaron en el Sheet; gana lo más reciente. */
  function aplicarCambiosSheet(data) {
    if (!data || !data.ok) return;
    const j = jornada(); if (!j) return;
    if (!state.progress[j.id]) state.progress[j.id] = {};
    const todas = data.filas || [];
    const coloreadas = todas.filter((f) => f.estado).length;
    const conDatos = todas.filter((f) => norm(f.fecha) || norm(f.observaciones) || norm(f.horasTexto)).length;
    let cambios = 0;
    (data.filas || []).forEach((f) => {
      const item = j.items.find((i) => i.serial === f.serial);
      if (!item || !f.estado || f.estado === 'PENDIENTE') return;
      // Si el Sheet solo tiene el color pero no datos, no sobrescribimos el equipo.
      if (!norm(f.fecha) && !norm(f.observaciones) && !norm(f.horasTexto)) return;
      const pr = state.progress[j.id][f.serial];
      if (pr && pr.fecha && norm(f.fecha) && norm(f.fecha) <= norm(pr.fecha)) return;
      const h = horasDeTexto(f.horasTexto);
      state.progress[j.id][f.serial] = {
        estado: f.estado,
        novedad: norm(f.observaciones) || observacionFinal(f.estado, '', item),
        horas: h || (pr?.horas || 1),
        horasTexto: norm(f.horasTexto) || horasTexto(pr?.horas || 1),
        fecha: norm(f.fecha) || (hoy() + ' ' + hora().slice(11)),
        tecnico: f.tecnico || state.config.tecnico || j.tecnico || '',
        // Un cambio real del Sheet (color + datos) también cuenta como registro.
        registrado: true,
      };
      item.estado = f.estado;
      item.novedad = state.progress[j.id][f.serial].novedad;
      cambios++;
    });
    if (cambios) {
      guardarProgreso(); guardarDataset();
      const activo = document.activeElement;
      const escribiendo = activo && /^(INPUT|TEXTAREA|SELECT)$/.test(activo.tagName);
      if (!escribiendo) render();
      toast(cambios + ' cambio(s) traídos del Sheet');
    }
    escribir(K.sync, { hora: hora(), ok: true, cambios: cambios, coloreadas: coloreadas, conDatos: conDatos });
  }

  let syncEnCurso = false;
  /** Enviar + traer. Se usa al abrir, cada 2 min, al volver la señal y al pulsar el botón. */
  async function sincronizar(silent) {
    const url = norm($('#cfgUrl')?.value || state.config.url);
    const out = $('#syncResult');
    if (!url) { if (!silent) toast('Configura la URL del Web App'); return false; }
    if (syncEnCurso) return false;
    syncEnCurso = true;
    const cfgUrl = url;
    const cfgToken = $('#cfgToken')?.value ?? state.config.token;
    const j = jornada(); const col = columnas();
    const filas = items().map((i) => filaPayload(i, progreso(j.id, i.serial), j, true));
    let ok = false, msg = '';
    try {
      await postSheet(cfgUrl, JSON.stringify({ token: cfgToken,
        hoja: nombreHoja(), jornada: j.id, fecha: hora(),
        tecnico: state.config.tecnico || j.tecnico || '', columnas: col,
        total: items().length, modo: 'total', filas }) );
      escribir(K.pendiente, []);
      ok = true; msg = 'Enviado y consultado ' + hora();
    } catch (e) { msg = 'No se pudo enviar: ' + e.message; }
    if (ok) leerDesdeSheets(cfgUrl, cfgToken);
    escribir(K.sync, { hora: hora(), ok: ok, cambios: 0 });
    if (!silent) {
      if (out) out.textContent = msg + '. Revisa la hoja ' + nombreHoja() + '.';
      if (ok) { toast('Datos enviados y consultados'); vibrar(100); setTimeout(() => render(), 2500); }
    }
    syncEnCurso = false;
    return ok;
  }

  async function sincronizarAuto() {
    if (!state.data) return;
    const pendiente = leer(K.cargar, null);
    if (pendiente && norm(state.config.url)) await subirCronograma(pendiente.hoja);
    if (!norm(state.config.url)) return;
    await reintentarPendientes();
    await sincronizar(true);
  }

  function generarEtiquetas() {
    const ed = $('#etiEd')?.value || ''; const dest = $('#etiquetas');
    const lista = items().filter((i) => !ed || i.edificio === ed).slice(0, 300);
    dest.innerHTML = '<div class="small muted">Generando…</div>';
    setTimeout(() => {
      dest.innerHTML = `<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:6px">
        ${lista.map((i) => `<div style="border:1px dashed #bbb;padding:5px;text-align:center">
          <svg class="barcode" data-code="${esc(i.serial)}"></svg>
          <div class="tiny mono">${esc(i.serial)}</div><div class="tiny muted">${esc(i.tipo || '')}</div></div>`).join('')}
      </div><p class="tiny muted mt">${lista.length} etiquetas. Ctrl+P para imprimir.</p>`;
      dest.querySelectorAll('svg.barcode').forEach((svg) => {
        try { JsBarcode(svg, svg.dataset.code, { format: 'CODE128', width: 1.4, height: 34, displayValue: false, margin: 0 }); } catch (_) {}
      });
    }, 30);
  }

  /* ---------------- eventos ---------------- */

  document.addEventListener('click', async (e) => {
    const nav = e.target.closest('#bottomnav button');
    if (nav) { go('#/' + nav.dataset.nav); return; }
    if (e.target.closest('#btnBack')) { history.back(); return; }
    if (e.target.closest('#btnHelp')) { go('#/ayuda'); return; }
    const b = e.target.closest('[data-act]'); if (!b) return;
    switch (b.dataset.act) {
      case 'go': go(b.dataset.hash); break;
      case 'area': go('#/area/' + encodeURIComponent(b.dataset.id)); break;
      case 'equipo': state.seleccion = null; state.seleccionSerial = null;
        go('#/equipo/' + encodeURIComponent(b.dataset.serial)); break;
      case 'sel': {
        state.seleccion = b.dataset.estado;
        view.querySelectorAll('.estado-btn').forEach((x) => x.classList.toggle('sel', x === b));
        const extra = $('#camposExtra'); if (!extra) break;
        extra.classList.toggle('hidden', !['NOVEDAD', 'AUSENTE', 'ALTA'].includes(state.seleccion));
        const lbl = extra.querySelector('label.campo');
        if (lbl) lbl.textContent = state.seleccion === 'NOVEDAD' ? '¿Qué novedad? (obligatorio)' : 'Observaciones (opcional)';
        if (state.seleccion === 'NOVEDAD' && !$('#foto')) {
          const l = document.createElement('label'); l.className = 'campo'; l.htmlFor = 'foto'; l.textContent = 'Foto (opcional)';
          const inp = document.createElement('input'); inp.className = 'campo'; inp.type = 'file'; inp.id = 'foto'; inp.accept = 'image/*'; inp.capture = 'environment';
          const d = document.createElement('div'); d.id = 'fotoActual'; d.className = 'mt';
          extra.append(l, inp, d);
          fotoLeer(state.seleccionSerial || '').then((url) => { if (url && $('#fotoActual')) $('#fotoActual').innerHTML = `<img src="${url}" style="max-width:100%;border-radius:10px">`; });
        }
        if (state.seleccion === 'ALTA' && !$('#destinoArea')) {
          const l = document.createElement('label'); l.className = 'campo'; l.htmlFor = 'destinoArea'; l.textContent = 'Dependencia donde está ahora';
          const sel = document.createElement('select'); sel.className = 'campo'; sel.id = 'destinoArea';
          sel.innerHTML = areas().map((a) => `<option value="${esc(a.id)}" ${a.id === state.areaActual ? 'selected' : ''}>${esc(a.oficina || a.piso)} · ${esc(a.piso)} · ${esc(a.edificio)}</option>`).join('');
          extra.insertBefore(sel, extra.querySelector('label.campo'));
          extra.insertBefore(l, sel);
        }
        break;
      }
      case 'horas': {
        const inp = $('#horasTxt'); if (inp) inp.value = horasTexto(Number(b.dataset.valor));
        view.querySelectorAll('[data-act=horas]').forEach((x) => x.classList.toggle('primary', x === b));
        break;
      }
      case 'guardar': await guardar(b.dataset.serial, b.dataset.siguiente === '1'); break;
      case 'confirmar-alta': darDeAlta(b.dataset.serial); break;
      case 'buscar-manual': buscarSerial($('#serialManual')?.value); break;
      case 'subir': importar('xlsx'); break;
      case 'elegir-hoja': await usarHoja(b.dataset.nombre); break;
      case 'informe': await generarInforme(); break;
      case 'exportar': exportar(); break;
      case 'sincronizar': await sincronizar(); break;
      case 'etiquetas': go('#/etiquetas'); break;
      case 'generar-etiquetas': generarEtiquetas(); break;
      case 'guardar-config':
        state.config = { url: $('#cfgUrl').value, token: $('#cfgToken').value, tecnico: $('#cfgTecnico').value };
        escribir(K.config, state.config); toast('Configuración guardada');
        await sincronizarAuto(); render(); break;
      case 'borrar-progreso':
        if (confirm('¿Borrar todos los registros de esta jornada en el teléfono?')) {
          delete state.progress[state.jornadaId]; guardarProgreso(); toast('Registros borrados'); render();
        } break;
      case 'olvidar':
        if (confirm('¿Olvidar el cronograma cargado?')) {
          [K.data, K.jornada, K.archivo].forEach((k) => localStorage.removeItem(k));
          location.reload();
        } break;
      case 'alta-ed': {
        const ed = b.value; const sel = $('#altaArea');
        if (sel) sel.innerHTML = areas().filter((a) => a.edificio === ed)
          .map((a) => `<option value="${esc(a.id)}">${esc(a.oficina || a.piso)} · ${esc(a.piso)}</option>`).join('');
        break;
      }
    }
  });

  view.addEventListener('input', (e) => {
    if (e.target.dataset.act === 'filtro-txt') {
      state.filtroTexto = e.target.value;
      clearTimeout(state.filtroTimer);
      state.filtroTimer = setTimeout(() => {
        const pos = e.target.selectionStart; render();
        const n = $('[data-act=filtro-txt]'); n?.focus(); n?.setSelectionRange(pos, pos);
      }, 350);
    }
  });
  view.addEventListener('change', (e) => {
    if (e.target.dataset.act === 'filtro-ed') { state.filtroEdificio = e.target.value; render(); }
  });
  window.addEventListener('hashchange', render);
  window.addEventListener('online', () => sincronizarAuto());
  document.addEventListener('visibilitychange', () => { if (!document.hidden) sincronizarAuto(); });
  setInterval(() => { if (!document.hidden) sincronizarAuto(); }, 2 * 60 * 1000);

  (async () => {
    await cargar(); render();
    setTimeout(() => sincronizarAuto(), 1500); // sincroniza al abrir la APK
    if ('serviceWorker' in navigator) window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(() => {}));
  })();
})();
