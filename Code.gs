/**
 * Gestión MP 2026 — Backend Apps Script
 * Hospital Hernán Henríquez Aravena · Temuco
 * Usuario único: Cristian, encargado biomédica HHHA.
 */

// ====================================================================
// CONSTANTES
// ====================================================================
const SHEETS = {
  PMP: 'PMP_2026',
  REG: 'Registro_MP-2026',
  EVENTOS: 'Eventos',
  PENDIENTES: 'Pendientes',
  REPROGS: 'Reprogramaciones',
  OVERRIDES: 'ResultadosOverride',
  ASIGNACIONES: 'Asignaciones',
  SNAPSHOT: 'Snapshot',
  INCONSISTENCIAS: 'Inconsistencias',
  CONFIG: 'Config',
};

const HEADERS = {
  EVENTOS: ['id','equipoKey','tipo','fecha','resultado','estadoEquipo','ejecutor','observacion','nEnvio','empresa','folio','comentario','folioGuia','createdAt'],
  PENDIENTES: ['id','equipoKey','descripcion','fechaCreacion','fechaCompromiso','ejecutor','prioridad','etiquetas','estado','tareas','actualizaciones'],
  REPROGS: ['id','equipoKey','mesOrigen','causal','mesDestino','fechaRegistro','comentario'],
  OVERRIDES: ['equipoKey','mes','resultado','fechaRegistro','usuario'],
  ASIGNACIONES: ['equipoKey','mes','ejecutor','fechaAsignacion'],
  SNAPSHOT: ['timestamp','equipoKey','mes','tipo','valor'],
  INCONSISTENCIAS: ['id','fechaDeteccion','equipoKey','tipo','mes','valorAntes','valorDespues','estado','comentario','pendienteVinculadoId'],
  CONFIG: ['key','value','descripcion'],
};

const MESES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const CAUSALES_CON_FECHA = ['C1','C5','C6','C7','C8'];
const CAUSALES_SIN_FECHA = ['C2','C3','C4'];
const RESULTADOS_VALIDOS = ['Si','Si-RA','C1','C2','C3','C4','C5','C6','C7','C8','FS','Baja','NU'];
const CODIGOS_PROG = ['X','R','RA','PM'];
const TIPOS_EVENTO = ['mp','envio','solicitud','recepcion','reparacion'];
const PRIORIDADES = ['alta','media','baja'];
const ESTADOS_PENDIENTE = ['abierto','cerrado'];
const ESTADOS_INCONSISTENCIA = ['nueva','revisada','descartada','convertida_pendiente'];

const EJECUTORES_DEFAULT = [
  'Carlos Bahamondes Seguel','Cristián Beltrán Oviedo','Cristina Rozas Urrutia',
  'Daniel Díaz Neira','Ignacio Berner Bergara','Macarena Toledo','Marco Ulloa',
  'Matías Soazo Garrido','Personal externo','Ricardo Matus Aroca','Tito Millapán Riquelme'
];

const CACHE_TTL_SEC = 300;
const SNAPSHOT_MAX = 5;

// ====================================================================
// API: doGet + include
// ====================================================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gestión MP 2026 — H.H.H.A.')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ====================================================================
// HELPERS
// ====================================================================
function ok(data) { return { ok: true, data: data }; }
function err(msg) { return { ok: false, error: String(msg) }; }
function wrap(fn) {
  return function() {
    try { return ok(fn.apply(null, arguments)); }
    catch (e) { return err(e && e.message ? e.message : e); }
  };
}

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name) {
  const s = ss().getSheetByName(name);
  if (!s) throw new Error('Hoja no encontrada: ' + name);
  return s;
}

function ensureSheet(name, headers) {
  let s = ss().getSheetByName(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    s.setFrozenRows(1);
    s.autoResizeColumns(1, headers.length);
  }
  return s;
}

function nowIso() { return Utilities.formatDate(new Date(), 'America/Santiago', "yyyy-MM-dd'T'HH:mm:ssXXX"); }
function todayIso() { return Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd'); }
function uid() { return Utilities.getUuid(); }
function userEmail() { try { return Session.getActiveUser().getEmail() || 'desconocido'; } catch (e) { return 'desconocido'; } }

function norm(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

function equipoKeyFromRow(idCol, invCol) {
  const inv = norm(invCol);
  if (inv && inv !== 'N/A') return 'inv:' + inv;
  return 'id:' + norm(idCol);
}

function familiaDe(nombreEquipo) {
  const e = String(nombreEquipo || '').toLowerCase();
  if (e.indexOf('monitor') >= 0) return 'Monitores';
  if (e.indexOf('ventilador') >= 0) return 'Ventiladores';
  if (e.indexOf('desfibrilador') >= 0 || e.indexOf('dea') >= 0) return 'Desfibriladores';
  if (e.indexOf('anestesia') >= 0) return 'M. Anestesia';
  if (e.indexOf('incubadora') >= 0) return 'Incubadoras';
  if (e.indexOf('diálisis') >= 0 || e.indexOf('dialisis') >= 0) return 'M. Diálisis';
  const extra = getConfigValue('FAMILIAS_EXTRA');
  if (extra) {
    const map = extra.split('|').map(function(s) { return s.trim(); }).filter(Boolean);
    for (let i = 0; i < map.length; i++) {
      const parts = map[i].split('=>');
      if (parts.length === 2 && e.indexOf(parts[0].trim().toLowerCase()) >= 0) return parts[1].trim();
    }
  }
  return 'Otros';
}

// ====================================================================
// CONFIG (lectura/escritura)
// ====================================================================
const _configMemo = {};

function getConfigValue(key) {
  if (_configMemo[key] !== undefined) return _configMemo[key];
  const s = ss().getSheetByName(SHEETS.CONFIG);
  if (!s) return '';
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) { _configMemo[key] = String(rows[i][1] || ''); return _configMemo[key]; }
  }
  return '';
}

function setConfigValue(key, value) {
  const s = ensureSheet(SHEETS.CONFIG, HEADERS.CONFIG);
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === key) {
      s.getRange(i + 1, 2).setValue(value);
      _configMemo[key] = String(value);
      return;
    }
  }
  s.appendRow([key, value, '']);
  _configMemo[key] = String(value);
}

const getConfig = wrap(function() {
  const s = ensureSheet(SHEETS.CONFIG, HEADERS.CONFIG);
  const rows = s.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) out[rows[i][0]] = { value: rows[i][1], descripcion: rows[i][2] || '' };
  }
  return out;
});

const setConfig = wrap(function(key, value) {
  setConfigValue(key, value);
  invalidateCache();
  return true;
});

// ====================================================================
// INIT
// ====================================================================
const inicializarHojas = wrap(function() {
  const created = [];

  ensureSheet(SHEETS.EVENTOS, HEADERS.EVENTOS);
  ensureSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES);
  ensureSheet(SHEETS.REPROGS, HEADERS.REPROGS);
  ensureSheet(SHEETS.OVERRIDES, HEADERS.OVERRIDES);
  ensureSheet(SHEETS.ASIGNACIONES, HEADERS.ASIGNACIONES);
  ensureSheet(SHEETS.SNAPSHOT, HEADERS.SNAPSHOT);
  ensureSheet(SHEETS.INCONSISTENCIAS, HEADERS.INCONSISTENCIAS);
  ensureSheet(SHEETS.CONFIG, HEADERS.CONFIG);

  // Validaciones dropdown
  const sEv = getSheet(SHEETS.EVENTOS);
  const rngEvTipo = sEv.getRange(2, 3, Math.max(1, sEv.getMaxRows() - 1), 1);
  rngEvTipo.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TIPOS_EVENTO, true).build());

  const sPe = getSheet(SHEETS.PENDIENTES);
  sPe.getRange(2, 7, Math.max(1, sPe.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PRIORIDADES, true).build());
  sPe.getRange(2, 9, Math.max(1, sPe.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_PENDIENTE, true).build());

  const sOv = getSheet(SHEETS.OVERRIDES);
  sOv.getRange(2, 3, Math.max(1, sOv.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RESULTADOS_VALIDOS, true).build());

  const sInc = getSheet(SHEETS.INCONSISTENCIAS);
  sInc.getRange(2, 8, Math.max(1, sInc.getMaxRows() - 1), 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS_INCONSISTENCIA, true).build());

  // Config inicial
  const sCfg = getSheet(SHEETS.CONFIG);
  const existing = {};
  const cRows = sCfg.getDataRange().getValues();
  for (let i = 1; i < cRows.length; i++) if (cRows[i][0]) existing[cRows[i][0]] = true;

  const defaults = [
    ['SPREADSHEET_ID', ss().getId(), 'ID del Google Sheet (autodetectado)'],
    ['WEB_APP_URL', '', 'URL del deployment de la Web App (pegar tras desplegar)'],
    ['EJECUTORES', EJECUTORES_DEFAULT.join('|'), 'Lista de ejecutores separada por |'],
    ['FAMILIAS_EXTRA', '', 'Mapeo extra "palabra=>Familia" separado por |'],
    ['ANIO_OPERATIVO', '2026', 'Año operativo activo'],
    ['TZ', 'America/Santiago', 'Zona horaria'],
    ['DEBUG', 'false', 'Modo debug (true/false)'],
    ['SNAPSHOT_AUTO', 'true', 'Snapshot automático al abrir si han pasado >24h'],
    ['ONEDIT_AUTO', 'false', 'Detección automática de cambios en Registro vía onEdit'],
    ['DARK_MODE_DEFAULT', 'false', 'Dark mode por defecto'],
  ];
  defaults.forEach(function(row) {
    if (!existing[row[0]]) { sCfg.appendRow(row); created.push(row[0]); }
  });

  // Asegurar que SPREADSHEET_ID siempre esté correcto
  setConfigValue('SPREADSHEET_ID', ss().getId());

  invalidateCache();
  return { created: created, spreadsheetId: ss().getId() };
});

// ====================================================================
// CACHE
// ====================================================================
function getCacheKey(k) { return 'mp2026_' + k; }

function invalidateCache() {
  try {
    const c = CacheService.getScriptCache();
    c.removeAll(['mp2026_equipos','mp2026_index']);
  } catch (e) {}
  for (const k in _configMemo) delete _configMemo[k];
}

// ====================================================================
// PARSE PMP / REGISTRO
// ====================================================================
function parseEquipos() {
  const sPmp = getSheet(SHEETS.PMP);
  const sReg = ss().getSheetByName(SHEETS.REG);
  const pmpData = sPmp.getRange(8, 1, Math.max(1, sPmp.getLastRow() - 7), 31).getValues();
  const regData = sReg
    ? sReg.getRange(8, 1, Math.max(1, sReg.getLastRow() - 7), 43).getValues()
    : [];

  const equipos = [];
  const invSeen = {};
  const dups = [];
  let filasReg = 0;

  for (let i = 0; i < pmpData.length; i++) {
    const row = pmpData[i] || [];
    const id = row[1];
    const inventario = norm(row[3]);
    const equipo = norm(row[4]);
    if ((id === null || id === '' || id === undefined) && !inventario && !equipo) continue;

    const mesesP = [];
    for (let m = 0; m < 12; m++) mesesP.push(norm(row[19 + m]));

    const regRow = regData[i] || [];
    const hasRegData = regRow.some(function(c) { return c !== null && c !== undefined && String(c).trim() !== ''; });
    if (hasRegData) filasReg++;

    const mesesReg = [];
    for (let m = 0; m < 12; m++) {
      mesesReg.push({ p: norm(regRow[19 + 2 * m]), r: norm(regRow[20 + 2 * m]) });
    }

    const isSlot = !equipo && !inventario;

    const eq = {
      id: String(id || ''),
      familia: norm(row[0]),
      carpeta: norm(row[2]),
      inventario: inventario,
      equipo: equipo,
      servicio: norm(row[5]),
      unidad: norm(row[6]),
      ubicacion: norm(row[7]),
      procedencia: norm(row[8]),
      marca: norm(row[9]),
      modelo: norm(row[10]),
      serie: norm(row[11]),
      anio: norm(row[12]),
      vidaUtil: norm(row[13]),
      clasificacion: norm(row[14]),
      enuBaja: norm(row[15]),
      frecuencia: norm(row[17]),
      mesesP: mesesP,
      mesesReg: mesesReg,
      isSlot: isSlot,
      familiaCalc: familiaDe(equipo),
    };
    eq.equipoKey = equipoKeyFromRow(id, inventario);
    equipos.push(eq);

    if (inventario && inventario !== 'N/A') {
      if (invSeen[inventario]) dups.push(inventario);
      else invSeen[inventario] = true;
    }
  }

  // Aplicar overrides + reprogramaciones + asignaciones
  const overrides = loadOverridesMap();
  const reprogs = loadReprogsMap();
  const asignaciones = loadAsignacionesMap();

  for (let i = 0; i < equipos.length; i++) {
    const eq = equipos[i];
    const ovr = overrides[eq.equipoKey] || {};
    const rep = reprogs[eq.equipoKey] || {};
    for (let m = 0; m < 12; m++) {
      if (ovr[m]) eq.mesesReg[m] = { p: eq.mesesReg[m].p, r: ovr[m] };
      if (rep[m] && !eq.mesesReg[m].p && !eq.mesesP[m]) eq.mesesReg[m].p = rep[m];
    }
    eq.asignaciones = asignaciones[eq.equipoKey] || {};
  }

  return {
    equipos: equipos,
    duplicados: Array.from(new Set(dups)),
    filasReg: filasReg,
    filasPMP: equipos.length,
  };
}

function loadOverridesMap() {
  const s = ss().getSheetByName(SHEETS.OVERRIDES);
  if (!s) return {};
  const data = s.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const k = data[i][0]; const mes = data[i][1]; const r = data[i][2];
    if (!k) continue;
    if (!out[k]) out[k] = {};
    out[k][Number(mes)] = String(r);
  }
  return out;
}

function loadReprogsMap() {
  const s = ss().getSheetByName(SHEETS.REPROGS);
  if (!s) return {};
  const data = s.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const k = data[i][1]; const dest = data[i][4];
    if (!k || dest === '' || dest === null) continue;
    if (!out[k]) out[k] = {};
    out[k][Number(dest) - 1] = 'R';
  }
  return out;
}

function loadAsignacionesMap() {
  const s = ss().getSheetByName(SHEETS.ASIGNACIONES);
  if (!s) return {};
  const data = s.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < data.length; i++) {
    const k = data[i][0]; const mes = data[i][1]; const ej = data[i][2];
    if (!k) continue;
    if (!out[k]) out[k] = {};
    out[k][Number(mes)] = String(ej);
  }
  return out;
}

// ====================================================================
// API: getEquipos / búsqueda
// ====================================================================
const getEquipos = wrap(function() {
  const c = CacheService.getScriptCache();
  const cached = c.get('mp2026_equipos');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const parsed = parseEquipos();
  try { c.put('mp2026_equipos', JSON.stringify(parsed), CACHE_TTL_SEC); } catch (e) {}
  return parsed;
});

function _norm(s) {
  return String(s || '').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

const buscarEquipos = wrap(function(query, filtros) {
  const data = parseEquipos();
  const equipos = data.equipos;
  query = String(query || '').trim();
  filtros = filtros || {};

  // Operadores: inv: serie: servicio: marca: modelo: ubic:
  const ops = {};
  const free = [];
  const tokens = query.split(/\s+/);
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    const m = t.match(/^(inv|serie|servicio|marca|modelo|ubic):(.+)$/i);
    if (m) ops[m[1].toLowerCase()] = _norm(m[2]);
    else if (t) free.push(_norm(t));
  }
  const freeQ = free.join(' ').trim();

  const out = equipos.filter(function(e) {
    if (e.isSlot) return false;
    if (filtros.servicio && e.servicio !== filtros.servicio) return false;
    if (filtros.familia && e.familiaCalc !== filtros.familia) return false;
    if (filtros.mesProg !== undefined && filtros.mesProg !== null && filtros.mesProg !== '') {
      const m = Number(filtros.mesProg);
      const cod = e.mesesP[m] || (e.mesesReg[m] && e.mesesReg[m].p);
      if (!cod || ['X','R','RA','PM'].indexOf(cod) < 0) return false;
    }
    if (filtros.responsable) {
      const tiene = Object.keys(e.asignaciones || {}).some(function(k) { return e.asignaciones[k] === filtros.responsable; });
      if (!tiene) return false;
    }
    if (filtros.causalActiva) {
      const tiene = e.mesesReg.some(function(c) { return c.r && c.r.charAt(0) === 'C'; });
      if (!tiene) return false;
    }
    if (ops.inv && _norm(e.inventario).indexOf(ops.inv) < 0) return false;
    if (ops.serie && _norm(e.serie).indexOf(ops.serie) < 0) return false;
    if (ops.servicio && _norm(e.servicio).indexOf(ops.servicio) < 0) return false;
    if (ops.marca && _norm(e.marca).indexOf(ops.marca) < 0) return false;
    if (ops.modelo && _norm(e.modelo).indexOf(ops.modelo) < 0) return false;
    if (ops.ubic && _norm(e.ubicacion).indexOf(ops.ubic) < 0) return false;
    if (freeQ) {
      const hay = [e.inventario, e.serie, e.equipo, e.marca, e.modelo, e.servicio, e.ubicacion].map(_norm).join(' ');
      if (hay.indexOf(freeQ) < 0) return false;
    }
    return true;
  });

  return { equipos: out, total: out.length, query: query };
});

// ====================================================================
// API: KPIs
// ====================================================================
const getKPIs = wrap(function(filtro) {
  const data = parseEquipos();
  const equipos = data.equipos.filter(function(e) { return !e.isSlot; });
  const mesActual = new Date().getMonth();

  let progAnio = 0, progMes = 0;
  let ejecAnio = 0, ejecMes = 0;
  let pendAnio = 0;
  let firmesAnio = 0;
  let causalCnt = {}; let resCntPorMes = {};
  let porMes = MESES.map(function(m) { return { mes: m, prog: 0, ejec: 0, pend: 0, firme: 0 }; });

  for (let i = 0; i < equipos.length; i++) {
    const eq = equipos[i];
    for (let m = 0; m < 12; m++) {
      const codP = eq.mesesP[m] || (eq.mesesReg[m] && eq.mesesReg[m].p);
      const codR = eq.mesesReg[m] && eq.mesesReg[m].r;
      if (codP && ['X','R','RA','PM'].indexOf(codP) >= 0) {
        progAnio++; porMes[m].prog++;
        if (m === mesActual) progMes++;
      }
      if (codR === 'Si' || codR === 'Si-RA') {
        ejecAnio++; porMes[m].ejec++;
        if (m === mesActual) ejecMes++;
      }
      if (codR === 'FS' || codR === 'Baja' || codR === 'NU') {
        firmesAnio++; porMes[m].firme++;
      }
      if (codR && codR.charAt(0) === 'C') {
        causalCnt[codR] = (causalCnt[codR] || 0) + 1;
        porMes[m].firme++;
      }
    }
  }

  pendAnio = Math.max(0, progAnio - ejecAnio - firmesAnio);
  for (let i = 0; i < porMes.length; i++) {
    porMes[i].pend = Math.max(0, porMes[i].prog - porMes[i].ejec - porMes[i].firme);
  }

  // Secundarios
  const eventosSheet = ss().getSheetByName(SHEETS.EVENTOS);
  let enST = 0, fueraServ = 0;
  if (eventosSheet) {
    const rows = eventosSheet.getDataRange().getValues();
    const ultimoPorEquipo = {};
    for (let i = 1; i < rows.length; i++) {
      const k = rows[i][1]; const fecha = rows[i][3]; const tipo = rows[i][2]; const estadoEquipo = rows[i][5];
      if (!k) continue;
      if (!ultimoPorEquipo[k] || (fecha && fecha > ultimoPorEquipo[k].fecha)) {
        ultimoPorEquipo[k] = { tipo: tipo, estado: estadoEquipo, fecha: fecha };
      }
    }
    for (const k in ultimoPorEquipo) {
      const u = ultimoPorEquipo[k];
      if (u.tipo === 'envio') enST++;
    }
  }
  equipos.forEach(function(eq) {
    if (eq.mesesReg.some(function(c) { return c.r === 'FS'; })) fueraServ++;
  });

  // Pendientes operativos vencidos
  let pendVencidos = 0;
  const pendSheet = ss().getSheetByName(SHEETS.PENDIENTES);
  if (pendSheet) {
    const rows = pendSheet.getDataRange().getValues();
    const hoy = todayIso();
    for (let i = 1; i < rows.length; i++) {
      const estado = rows[i][8]; const fc = rows[i][4];
      if (estado === 'abierto' && fc && String(fc) < hoy) pendVencidos++;
    }
  }

  // Inconsistencias sin revisar
  let incSinRevisar = 0;
  const incSheet = ss().getSheetByName(SHEETS.INCONSISTENCIAS);
  if (incSheet) {
    const rows = incSheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][7] === 'nueva') incSinRevisar++;
    }
  }

  const cumplimiento = progAnio > 0 ? Math.round((ejecAnio / progAnio) * 100) : 0;
  return {
    principales: {
      programadas: { anio: progAnio, mes: progMes },
      ejecutadas: { anio: ejecAnio, mes: ejecMes, pct: cumplimiento },
      pendientes: { anio: pendAnio },
    },
    secundarios: {
      enServicioTecnico: enST,
      fueraServicio: fueraServ,
      pendientesVencidos: pendVencidos,
      inconsistenciasSinRevisar: incSinRevisar,
    },
    porMes: porMes,
    causales: causalCnt,
    mesActual: mesActual,
  };
});

// ====================================================================
// API: Pendientes CRUD
// ====================================================================
function rowsToObjects(sheet, headers) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].every(function(c) { return c === '' || c === null; })) continue;
    const o = {};
    for (let j = 0; j < headers.length; j++) o[headers[j]] = data[i][j];
    out.push(o);
  }
  return out;
}

function findRow(sheet, col, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]) === String(value)) return i + 1;
  }
  return -1;
}

const getPendientes = wrap(function(filtro) {
  const s = ensureSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES);
  const list = rowsToObjects(s, HEADERS.PENDIENTES).map(function(p) {
    try { p.tareas = p.tareas ? JSON.parse(p.tareas) : []; } catch (e) { p.tareas = []; }
    try { p.actualizaciones = p.actualizaciones ? JSON.parse(p.actualizaciones) : []; } catch (e) { p.actualizaciones = []; }
    p.etiquetas = p.etiquetas ? String(p.etiquetas).split('|').map(function(x) { return x.trim(); }).filter(Boolean) : [];
    return p;
  });
  filtro = filtro || {};
  const hoy = todayIso();
  return list.filter(function(p) {
    if (filtro.equipoKey && p.equipoKey !== filtro.equipoKey) return false;
    if (filtro.estado && p.estado !== filtro.estado) return false;
    if (filtro.ejecutor && p.ejecutor !== filtro.ejecutor) return false;
    if (filtro.prioridad && p.prioridad !== filtro.prioridad) return false;
    if (filtro.vencidos && !(p.estado === 'abierto' && p.fechaCompromiso && String(p.fechaCompromiso) < hoy)) return false;
    return true;
  });
});

const savePendiente = wrap(function(pend) {
  const s = ensureSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES);
  if (!pend.descripcion) throw new Error('Descripción requerida');
  pend.id = pend.id || uid();
  pend.fechaCreacion = pend.fechaCreacion || todayIso();
  pend.estado = pend.estado || 'abierto';
  pend.prioridad = pend.prioridad || 'media';
  pend.tareas = JSON.stringify(pend.tareas || []);
  pend.actualizaciones = JSON.stringify(pend.actualizaciones || [{ id: uid(), fecha: todayIso(), texto: 'Pendiente creado' }]);
  pend.etiquetas = Array.isArray(pend.etiquetas) ? pend.etiquetas.join('|') : String(pend.etiquetas || '');

  const row = HEADERS.PENDIENTES.map(function(h) { return pend[h] !== undefined ? pend[h] : ''; });
  const existing = findRow(s, 0, pend.id);
  if (existing > 0) s.getRange(existing, 1, 1, row.length).setValues([row]);
  else s.appendRow(row);
  return pend.id;
});

const togglePendiente = wrap(function(id) {
  const s = ensureSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES);
  const r = findRow(s, 0, id);
  if (r < 0) throw new Error('Pendiente no encontrado');
  const estado = s.getRange(r, 9).getValue();
  const nuevo = estado === 'abierto' ? 'cerrado' : 'abierto';
  s.getRange(r, 9).setValue(nuevo);
  const actCell = s.getRange(r, 11);
  let acts = []; try { acts = JSON.parse(actCell.getValue() || '[]'); } catch (e) { acts = []; }
  acts.push({ id: uid(), fecha: todayIso(), texto: nuevo === 'cerrado' ? 'Pendiente cerrado' : 'Pendiente reabierto' });
  actCell.setValue(JSON.stringify(acts));
  return nuevo;
});

const deletePendiente = wrap(function(id) {
  const s = ensureSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES);
  const r = findRow(s, 0, id);
  if (r < 0) throw new Error('Pendiente no encontrado');
  s.deleteRow(r);
  return true;
});

// ====================================================================
// API: Eventos CRUD
// ====================================================================
const getEventos = wrap(function(equipoKey) {
  const s = ensureSheet(SHEETS.EVENTOS, HEADERS.EVENTOS);
  return rowsToObjects(s, HEADERS.EVENTOS).filter(function(e) {
    return !equipoKey || e.equipoKey === equipoKey;
  });
});

const saveEvento = wrap(function(ev) {
  const s = ensureSheet(SHEETS.EVENTOS, HEADERS.EVENTOS);
  if (!ev.equipoKey) throw new Error('equipoKey requerido');
  if (!ev.tipo || TIPOS_EVENTO.indexOf(ev.tipo) < 0) throw new Error('Tipo inválido');
  if (!ev.fecha) throw new Error('Fecha requerida');
  ev.id = ev.id || uid();
  ev.createdAt = ev.createdAt || nowIso();
  const row = HEADERS.EVENTOS.map(function(h) { return ev[h] !== undefined ? ev[h] : ''; });
  const r = findRow(s, 0, ev.id);
  if (r > 0) s.getRange(r, 1, 1, row.length).setValues([row]);
  else s.appendRow(row);

  // Si es MP con resultado: registrar override en el mes de la fecha
  if (ev.tipo === 'mp' && ev.resultado) {
    const mes = (new Date(ev.fecha)).getMonth();
    saveResultadoOverrideInternal(ev.equipoKey, mes, ev.resultado);
    // Si causal con fecha y mesDestino, registrar reprogramación
    if (CAUSALES_CON_FECHA.indexOf(ev.resultado) >= 0 && ev.mesDestino !== '' && ev.mesDestino !== undefined && ev.mesDestino !== null) {
      saveReprogramacionInternal(ev.equipoKey, mes + 1, ev.resultado, Number(ev.mesDestino) + 1, ev.observacion || '');
    }
  }

  invalidateCache();
  return ev.id;
});

const deleteEvento = wrap(function(id) {
  const s = ensureSheet(SHEETS.EVENTOS, HEADERS.EVENTOS);
  const r = findRow(s, 0, id);
  if (r < 0) throw new Error('Evento no encontrado');
  s.deleteRow(r);
  invalidateCache();
  return true;
});

// ====================================================================
// API: Reprogramaciones / Overrides / Asignaciones
// ====================================================================
function saveReprogramacionInternal(equipoKey, mesOrigen, causal, mesDestino, comentario) {
  const s = ensureSheet(SHEETS.REPROGS, HEADERS.REPROGS);
  const row = [uid(), equipoKey, mesOrigen, causal, mesDestino || '', todayIso(), comentario || ''];
  s.appendRow(row);
}

const saveReprogramacion = wrap(function(equipoKey, mesOrigen, causal, mesDestino, comentario) {
  if (!equipoKey) throw new Error('equipoKey requerido');
  saveReprogramacionInternal(equipoKey, mesOrigen, causal, mesDestino, comentario);
  invalidateCache();
  return true;
});

function saveResultadoOverrideInternal(equipoKey, mes, resultado) {
  const s = ensureSheet(SHEETS.OVERRIDES, HEADERS.OVERRIDES);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === equipoKey && Number(data[i][1]) === Number(mes)) {
      s.getRange(i + 1, 3).setValue(resultado);
      s.getRange(i + 1, 4).setValue(nowIso());
      s.getRange(i + 1, 5).setValue(userEmail());
      return;
    }
  }
  s.appendRow([equipoKey, mes, resultado, nowIso(), userEmail()]);
}

const saveResultadoOverride = wrap(function(equipoKey, mes, resultado) {
  if (!equipoKey) throw new Error('equipoKey requerido');
  if (RESULTADOS_VALIDOS.indexOf(resultado) < 0 && resultado !== '') throw new Error('Resultado inválido');
  saveResultadoOverrideInternal(equipoKey, Number(mes), resultado);
  invalidateCache();
  return true;
});

const saveAsignacion = wrap(function(equipoKey, mes, ejecutor) {
  const s = ensureSheet(SHEETS.ASIGNACIONES, HEADERS.ASIGNACIONES);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === equipoKey && Number(data[i][1]) === Number(mes)) {
      s.getRange(i + 1, 3).setValue(ejecutor);
      s.getRange(i + 1, 4).setValue(nowIso());
      invalidateCache();
      return true;
    }
  }
  s.appendRow([equipoKey, Number(mes), ejecutor, nowIso()]);
  invalidateCache();
  return true;
});

// ====================================================================
// API: Plantilla mensual
// ====================================================================
const generarPlantillaMensual = wrap(function(mes) {
  mes = Number(mes);
  if (mes < 0 || mes > 11) throw new Error('Mes inválido');
  const data = parseEquipos();
  const equiposMes = data.equipos.filter(function(eq) {
    if (eq.isSlot) return false;
    const c = eq.mesesP[mes] || (eq.mesesReg[mes] && eq.mesesReg[mes].p);
    return ['X','R','RA','PM'].indexOf(c) >= 0;
  });

  const nombre = 'Plantilla_' + MESES[mes] + '_2026';
  let sheet = ss().getSheetByName(nombre);
  if (sheet) ss().deleteSheet(sheet);
  sheet = ss().insertSheet(nombre);

  const headers = ['N° Carpeta','N° Inventario','Equipo','Servicio','Unidad','Ubicación','Marca','Modelo','Serie','Año','Frecuencia','Programado','Responsable'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1e40af').setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  const rows = equiposMes.map(function(eq) {
    const resp = (eq.asignaciones && eq.asignaciones[mes]) || '';
    const cod = eq.mesesP[mes] || (eq.mesesReg[mes] && eq.mesesReg[mes].p) || '';
    return [eq.carpeta, eq.inventario, eq.equipo, eq.servicio, eq.unidad, eq.ubicacion, eq.marca, eq.modelo, eq.serie, eq.anio, eq.frecuencia, cod, resp];
  });
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const ejecutores = (getConfigValue('EJECUTORES') || EJECUTORES_DEFAULT.join('|')).split('|').filter(Boolean);
  const validation = SpreadsheetApp.newDataValidation().requireValueInList(ejecutores, true).setAllowInvalid(false).build();
  if (rows.length) sheet.getRange(2, 13, rows.length, 1).setDataValidation(validation);

  sheet.autoResizeColumns(1, headers.length);
  return {
    url: ss().getUrl() + '#gid=' + sheet.getSheetId(),
    count: equiposMes.length,
    nombre: nombre,
    gid: sheet.getSheetId(),
  };
});

const aplicarPlantillaCargada = wrap(function(mes) {
  mes = Number(mes);
  const nombre = 'Plantilla_' + MESES[mes] + '_2026';
  const sheet = ss().getSheetByName(nombre);
  if (!sheet) throw new Error('Plantilla no encontrada: ' + nombre);
  const data = sheet.getDataRange().getValues();
  let asignados = 0, sinResp = 0;
  const noEncontrados = [];
  const parsed = parseEquipos();
  const byInv = {}; parsed.equipos.forEach(function(e) { if (e.inventario) byInv[e.inventario] = e; });

  for (let i = 1; i < data.length; i++) {
    const inv = String(data[i][1] || '').trim();
    const resp = String(data[i][12] || '').trim();
    if (!inv) continue;
    if (!resp) { sinResp++; continue; }
    const eq = byInv[inv];
    if (!eq) { noEncontrados.push(inv); continue; }
    const r = saveAsignacion(eq.equipoKey, mes, resp);
    if (r.ok) asignados++;
  }
  invalidateCache();
  return { asignados: asignados, sinResp: sinResp, noEncontrados: noEncontrados };
});

// ====================================================================
// API: Snapshot + diff (Inconsistencias)
// ====================================================================
function loadLastSnapshot() {
  const s = ss().getSheetByName(SHEETS.SNAPSHOT);
  if (!s || s.getLastRow() < 2) return null;
  const data = s.getDataRange().getValues();
  // último timestamp = max
  let lastTs = '';
  for (let i = 1; i < data.length; i++) if (data[i][0] > lastTs) lastTs = data[i][0];
  if (!lastTs) return null;
  const snap = {}; const snapEquipos = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== lastTs) continue;
    const k = data[i][1]; const m = Number(data[i][2]); const tipo = data[i][3]; const v = data[i][4];
    if (!snap[k]) snap[k] = { mesesP: Array(12).fill(''), mesesR: Array(12).fill('') };
    if (tipo === 'P') snap[k].mesesP[m] = v;
    if (tipo === 'R') snap[k].mesesR[m] = v;
    if (tipo === 'CAT') snapEquipos[k] = (function() { try { return JSON.parse(v); } catch (e) { return null; } })();
  }
  return { timestamp: lastTs, equipos: snap, catastro: snapEquipos };
}

const tomarSnapshot = wrap(function() {
  const data = parseEquipos();
  const s = ensureSheet(SHEETS.SNAPSHOT, HEADERS.SNAPSHOT);
  const ts = nowIso();
  const rows = [];
  data.equipos.forEach(function(eq) {
    if (eq.isSlot) return;
    for (let m = 0; m < 12; m++) {
      const p = eq.mesesReg[m].p || eq.mesesP[m];
      const r = eq.mesesReg[m].r;
      if (p) rows.push([ts, eq.equipoKey, m, 'P', p]);
      if (r) rows.push([ts, eq.equipoKey, m, 'R', r]);
    }
    // Catastro: una fila por equipo
    rows.push([ts, eq.equipoKey, -1, 'CAT', JSON.stringify({ inv: eq.inventario, eq: eq.equipo, serv: eq.servicio, ubic: eq.ubicacion, marca: eq.marca, modelo: eq.modelo })]);
  });
  if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, 5).setValues(rows);

  purgeOldSnapshots(s);
  setConfigValue('LAST_SNAPSHOT', ts);
  return { timestamp: ts, filas: rows.length };
});

function purgeOldSnapshots(s) {
  const data = s.getDataRange().getValues();
  const tsSet = {}; const tsList = [];
  for (let i = 1; i < data.length; i++) if (data[i][0]) { if (!tsSet[data[i][0]]) { tsSet[data[i][0]] = true; tsList.push(data[i][0]); } }
  tsList.sort();
  const toKeep = tsList.slice(-SNAPSHOT_MAX).reduce(function(acc, t) { acc[t] = true; return acc; }, {});
  const newRows = [HEADERS.SNAPSHOT];
  for (let i = 1; i < data.length; i++) if (toKeep[data[i][0]]) newRows.push(data[i]);
  s.clearContents();
  s.getRange(1, 1, newRows.length, 5).setValues(newRows);
}

const compararSnapshot = wrap(function() {
  const snap = loadLastSnapshot();
  if (!snap) throw new Error('No hay snapshots previos. Tome uno primero.');
  const data = parseEquipos();
  const incSheet = ensureSheet(SHEETS.INCONSISTENCIAS, HEADERS.INCONSISTENCIAS);

  const ahora = {}; const ahoraCat = {};
  data.equipos.forEach(function(eq) {
    if (eq.isSlot) return;
    ahora[eq.equipoKey] = {
      mesesP: eq.mesesReg.map(function(c, i) { return c.p || eq.mesesP[i] || ''; }),
      mesesR: eq.mesesReg.map(function(c) { return c.r || ''; }),
    };
    ahoraCat[eq.equipoKey] = { inv: eq.inventario, eq: eq.equipo, serv: eq.servicio, ubic: eq.ubicacion, marca: eq.marca, modelo: eq.modelo };
  });

  const overrides = loadOverridesMap();
  const detectadas = [];
  const fecha = nowIso();

  // 1-4, 8: por equipo presente en ambos
  for (const k in snap.equipos) {
    const prev = snap.equipos[k]; const cur = ahora[k];
    if (!cur) { detectadas.push([uid(), fecha, k, 'equipo_desaparecido', '', '', '', 'nueva', '', '']); continue; }
    for (let m = 0; m < 12; m++) {
      const prevR = prev.mesesR[m]; const curR = cur.mesesR[m];
      if (prevR && (prevR === 'Si' || prevR === 'Si-RA') && (!curR || curR === '')) {
        detectadas.push([uid(), fecha, k, 'si_desaparecida', m, prevR, curR, 'nueva', '', '']);
      }
      else if (prevR === 'Si' && curR && (curR.charAt(0) === 'C' || curR === 'FS' || curR === 'Baja' || curR === 'NU')) {
        detectadas.push([uid(), fecha, k, 'si_a_causal', m, prevR, curR, 'nueva', '', '']);
      }
      else if (prevR && prevR.charAt(0) === 'C' && curR && curR.charAt(0) === 'C' && prevR !== curR) {
        detectadas.push([uid(), fecha, k, 'causal_cambiada', m, prevR, curR, 'nueva', '', '']);
      }
      if (curR === 'Baja' && prevR !== 'Baja') {
        detectadas.push([uid(), fecha, k, 'baja_nueva', m, prevR || '', curR, 'nueva', '', '']);
      }
      // override no reflejado
      if (overrides[k] && overrides[k][m] && overrides[k][m] === 'Si' && curR !== 'Si') {
        detectadas.push([uid(), fecha, k, 'override_no_reflejado', m, 'Si (override)', curR || '', 'nueva', '', '']);
      }
    }
    // cambio_catastro
    const prevCat = snap.catastro && snap.catastro[k]; const curCat = ahoraCat[k];
    if (prevCat && curCat) {
      ['serv','ubic','marca','modelo'].forEach(function(field) {
        if (prevCat[field] && curCat[field] && prevCat[field] !== curCat[field]) {
          detectadas.push([uid(), fecha, k, 'cambio_catastro', '', field + ':' + prevCat[field], field + ':' + curCat[field], 'nueva', '', '']);
        }
      });
    }
  }
  // 6: equipo_nuevo
  for (const k in ahora) {
    if (!snap.equipos[k]) detectadas.push([uid(), fecha, k, 'equipo_nuevo', '', '', '', 'nueva', '', '']);
  }

  if (detectadas.length) incSheet.getRange(incSheet.getLastRow() + 1, 1, detectadas.length, HEADERS.INCONSISTENCIAS.length).setValues(detectadas);

  invalidateCache();
  return { detectadas: detectadas.length, snapshot: snap.timestamp };
});

const getInconsistencias = wrap(function(filtro) {
  const s = ensureSheet(SHEETS.INCONSISTENCIAS, HEADERS.INCONSISTENCIAS);
  const list = rowsToObjects(s, HEADERS.INCONSISTENCIAS);
  filtro = filtro || {};
  return list.filter(function(it) {
    if (filtro.estado && it.estado !== filtro.estado) return false;
    if (filtro.tipo && it.tipo !== filtro.tipo) return false;
    if (filtro.equipoKey && it.equipoKey !== filtro.equipoKey) return false;
    return true;
  });
});

const marcarInconsistencia = wrap(function(id, estado, comentario) {
  const s = ensureSheet(SHEETS.INCONSISTENCIAS, HEADERS.INCONSISTENCIAS);
  const r = findRow(s, 0, id);
  if (r < 0) throw new Error('Inconsistencia no encontrada');
  s.getRange(r, 8).setValue(estado);
  if (comentario !== undefined) s.getRange(r, 9).setValue(comentario);
  return true;
});

const convertirInconsistenciaEnPendiente = wrap(function(id) {
  const s = ensureSheet(SHEETS.INCONSISTENCIAS, HEADERS.INCONSISTENCIAS);
  const r = findRow(s, 0, id);
  if (r < 0) throw new Error('Inconsistencia no encontrada');
  const row = s.getRange(r, 1, 1, HEADERS.INCONSISTENCIAS.length).getValues()[0];
  const equipoKey = row[2]; const tipo = row[3]; const mes = row[4];
  const valorAntes = row[5]; const valorDespues = row[6];
  const desc = 'Revisar inconsistencia: ' + tipo + (mes !== '' && mes !== null ? ' (mes ' + MESES[Number(mes)] + ')' : '') +
    ' · antes: "' + valorAntes + '" · ahora: "' + valorDespues + '"';
  const pendId = uid();
  const pend = {
    id: pendId, equipoKey: equipoKey, descripcion: desc,
    fechaCreacion: todayIso(),
    fechaCompromiso: Utilities.formatDate(new Date(Date.now() + 7 * 86400000), 'America/Santiago', 'yyyy-MM-dd'),
    ejecutor: '', prioridad: 'media', etiquetas: 'inconsistencia',
    estado: 'abierto',
    tareas: [],
    actualizaciones: [{ id: uid(), fecha: todayIso(), texto: 'Generado desde inconsistencia ' + id }],
  };
  const sp = savePendiente(pend);
  s.getRange(r, 8).setValue('convertida_pendiente');
  s.getRange(r, 10).setValue(pendId);
  return { pendienteId: pendId };
});

// ====================================================================
// API: Verificación carga
// ====================================================================
const getVerificacionCarga = wrap(function() {
  const data = parseEquipos();
  const validos = data.equipos.filter(function(e) { return !e.isSlot; }).length;
  const slots = data.equipos.length - validos;

  const FAMILIAS_ORDEN = ['Monitores','Ventiladores','Desfibriladores','M. Anestesia','Incubadoras','M. Diálisis','Otros'];
  const famCount = {}; FAMILIAS_ORDEN.forEach(function(f) { famCount[f] = 0; });
  data.equipos.forEach(function(eq) {
    if (eq.isSlot) return;
    const f = eq.familiaCalc; famCount[f] = (famCount[f] || 0) + 1;
  });

  const pmpCounts = {}; CODIGOS_PROG.forEach(function(c) { pmpCounts[c] = Array(12).fill(0); });
  const regCounts = {}; RESULTADOS_VALIDOS.forEach(function(c) { regCounts[c] = Array(12).fill(0); });

  data.equipos.forEach(function(eq) {
    for (let m = 0; m < 12; m++) {
      const p = eq.mesesP[m] || eq.mesesReg[m].p;
      if (p) { if (!pmpCounts[p]) pmpCounts[p] = Array(12).fill(0); pmpCounts[p][m]++; }
      const r = eq.mesesReg[m].r;
      if (r) { if (!regCounts[r]) regCounts[r] = Array(12).fill(0); regCounts[r][m]++; }
    }
  });

  const slotsConProg = data.equipos.filter(function(eq) {
    return eq.isSlot && (eq.mesesP.some(function(v) { return v; }) || eq.mesesReg.some(function(c) { return c.p || c.r; }));
  }).length;

  return {
    filasPMP: data.filasPMP, filasReg: data.filasReg,
    validos: validos, slots: slots,
    famCount: famCount, pmpCounts: pmpCounts, regCounts: regCounts,
    duplicados: data.duplicados, slotsConProg: slotsConProg,
    cargadoEn: nowIso(),
  };
});

// ====================================================================
// API: Export / Import respaldo JSON
// ====================================================================
const exportarRespaldo = wrap(function() {
  const sheets = [SHEETS.EVENTOS, SHEETS.PENDIENTES, SHEETS.REPROGS, SHEETS.OVERRIDES, SHEETS.ASIGNACIONES, SHEETS.INCONSISTENCIAS];
  const out = { ts: nowIso(), version: 'v9' };
  sheets.forEach(function(name) {
    const s = ss().getSheetByName(name);
    if (!s) return;
    const data = s.getDataRange().getValues();
    out[name] = data;
  });
  // catastro
  out.equipos = parseEquipos().equipos;
  return out;
});

const importarRespaldo = wrap(function(json) {
  if (typeof json === 'string') json = JSON.parse(json);
  if (!json) throw new Error('Respaldo vacío');

  // Detectar formato HTML v4 (keys: eventos, pendientes, reprogs, resultadosOverride, asignaciones)
  if (json.eventos || json.pendientes || json.reprogs || json.resultadosOverride || json.asignaciones) {
    return importarRespaldoHtml(json);
  }
  // Formato GAS (matrices por hoja)
  if (json[SHEETS.EVENTOS]) {
    rewriteSheet(SHEETS.EVENTOS, HEADERS.EVENTOS, json[SHEETS.EVENTOS]);
    rewriteSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES, json[SHEETS.PENDIENTES]);
    rewriteSheet(SHEETS.REPROGS, HEADERS.REPROGS, json[SHEETS.REPROGS]);
    rewriteSheet(SHEETS.OVERRIDES, HEADERS.OVERRIDES, json[SHEETS.OVERRIDES]);
    rewriteSheet(SHEETS.ASIGNACIONES, HEADERS.ASIGNACIONES, json[SHEETS.ASIGNACIONES]);
    rewriteSheet(SHEETS.INCONSISTENCIAS, HEADERS.INCONSISTENCIAS, json[SHEETS.INCONSISTENCIAS]);
  }
  invalidateCache();
  return { ok: true };
});

function rewriteSheet(name, headers, data) {
  if (!data) return;
  const s = ensureSheet(name, headers);
  s.clearContents();
  s.getRange(1, 1, data.length || 1, headers.length).setValues(data.length ? data : [headers]);
}

function importarRespaldoHtml(json) {
  // Eventos
  const sEv = ensureSheet(SHEETS.EVENTOS, HEADERS.EVENTOS);
  const evRows = [];
  if (json.eventos) {
    for (const k in json.eventos) {
      (json.eventos[k] || []).forEach(function(ev) {
        evRows.push([
          ev.id || uid(), k, ev.tipo || 'mp', ev.fecha || '',
          ev.resultado || '', ev.estadoEquipo || '', ev.ejecutor || '',
          ev.observacion || '', ev.numero || '', ev.empresa || '',
          ev.folio || '', ev.comentario || '', ev.folioGuia || '',
          nowIso(),
        ]);
      });
    }
  }
  if (evRows.length) sEv.getRange(sEv.getLastRow() + 1, 1, evRows.length, HEADERS.EVENTOS.length).setValues(evRows);

  // Pendientes
  const sPe = ensureSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES);
  const peRows = [];
  if (json.pendientes) {
    for (const k in json.pendientes) {
      (json.pendientes[k] || []).forEach(function(p) {
        peRows.push([
          p.id || uid(), k, p.descripcion || '', p.fecha || todayIso(),
          p.fechaCompromiso || '', p.ejecutor || '', p.prioridad || 'media',
          Array.isArray(p.etiquetas) ? p.etiquetas.join('|') : (p.etiquetas || ''),
          p.estado || 'abierto',
          JSON.stringify(p.tareas || []),
          JSON.stringify(p.actualizaciones || []),
        ]);
      });
    }
  }
  if (peRows.length) sPe.getRange(sPe.getLastRow() + 1, 1, peRows.length, HEADERS.PENDIENTES.length).setValues(peRows);

  // Reprogs
  const sRe = ensureSheet(SHEETS.REPROGS, HEADERS.REPROGS);
  const reRows = [];
  if (json.reprogs) {
    for (const k in json.reprogs) {
      const map = json.reprogs[k] || {};
      for (const m in map) {
        reRows.push([uid(), k, '', map[m], Number(m) + 1, todayIso(), 'importado HTML']);
      }
    }
  }
  if (reRows.length) sRe.getRange(sRe.getLastRow() + 1, 1, reRows.length, HEADERS.REPROGS.length).setValues(reRows);

  // Overrides
  const sOv = ensureSheet(SHEETS.OVERRIDES, HEADERS.OVERRIDES);
  const ovRows = [];
  if (json.resultadosOverride) {
    for (const k in json.resultadosOverride) {
      const map = json.resultadosOverride[k] || {};
      for (const m in map) ovRows.push([k, Number(m), map[m], nowIso(), userEmail()]);
    }
  }
  if (ovRows.length) sOv.getRange(sOv.getLastRow() + 1, 1, ovRows.length, HEADERS.OVERRIDES.length).setValues(ovRows);

  // Asignaciones
  const sAs = ensureSheet(SHEETS.ASIGNACIONES, HEADERS.ASIGNACIONES);
  const asRows = [];
  if (json.asignaciones) {
    for (const k in json.asignaciones) {
      const map = json.asignaciones[k] || {};
      for (const m in map) asRows.push([k, Number(m), map[m], nowIso()]);
    }
  }
  if (asRows.length) sAs.getRange(sAs.getLastRow() + 1, 1, asRows.length, HEADERS.ASIGNACIONES.length).setValues(asRows);

  invalidateCache();
  return { eventos: evRows.length, pendientes: peRows.length, reprogs: reRows.length, overrides: ovRows.length, asignaciones: asRows.length };
}

// ====================================================================
// API: bootstrap (devuelve todo lo que el cliente necesita al cargar)
// ====================================================================
const bootstrap = wrap(function() {
  // Asegurar hojas iniciales (idempotente)
  try { inicializarHojas(); } catch (e) {}
  const cfgRes = getConfig();
  const cfg = cfgRes.ok ? cfgRes.data : {};
  const ejecutores = (cfg.EJECUTORES && cfg.EJECUTORES.value ? String(cfg.EJECUTORES.value) : EJECUTORES_DEFAULT.join('|')).split('|').filter(Boolean);
  const lastSnapshot = cfg.LAST_SNAPSHOT && cfg.LAST_SNAPSHOT.value ? String(cfg.LAST_SNAPSHOT.value) : '';

  // Snapshot automático si > 24h
  if (cfg.SNAPSHOT_AUTO && String(cfg.SNAPSHOT_AUTO.value) === 'true') {
    if (!lastSnapshot || (Date.now() - new Date(lastSnapshot).getTime()) > 24 * 3600 * 1000) {
      try { tomarSnapshot(); } catch (e) {}
    }
  }

  return {
    config: cfg,
    ejecutores: ejecutores,
    spreadsheetId: ss().getId(),
    spreadsheetUrl: ss().getUrl(),
    userEmail: userEmail(),
    meses: MESES,
    causales: {
      C1: 'Imposibilidad de desocupar el equipo del paciente por indicación clínica',
      C2: 'Equipo en servicio técnico',
      C3: 'Equipo no operativo, a la espera de repuestos o accesorios',
      C4: 'Equipo en préstamo a otro hospital o institución',
      C5: 'No disponibilidad de horas hombre del funcionario SEC por alta carga laboral',
      C6: 'No disponibilidad de horas hombre del servicio técnico externo',
      C7: 'Ausencia justificada del funcionario SEC superior a 15 días',
      C8: 'Contingencia hospitalaria',
    },
    resultados: RESULTADOS_VALIDOS,
    codigosProg: CODIGOS_PROG,
    tiposEvento: TIPOS_EVENTO,
  };
});

// ====================================================================
// API: Datos sintéticos
// ====================================================================
const seedTestData = wrap(function() {
  // Crear o limpiar hojas de PMP y Registro de prueba
  const familias = ['Monitor Multiparametro','Ventilador mecánico','Desfibrilador','Máquina de Anestesia','Incubadora neonatal','Máquina de Diálisis','Bomba de infusión'];
  const servicios = ['UCI Adultos','UCI Pediátrica','UCI Neonatal','Pabellón','Urgencias','Hemodiálisis','Medicina Interna'];
  const ubicaciones = ['Sala 1','Sala 2','Sala 3','Box 1','Box 2','Pasillo','Bodega'];
  const marcas = ['Philips','Mindray','Drager','GE','Siemens','Stryker','BAXTER'];

  const sPmp = ss().getSheetByName(SHEETS.PMP) || ss().insertSheet(SHEETS.PMP);
  const sReg = ss().getSheetByName(SHEETS.REG) || ss().insertSheet(SHEETS.REG);

  const headersPmp = ['Familia','ID','N°Carpeta','N°Inventario','Equipo','Servicio','Unidad','Ubicación','Procedencia','Marca','Modelo','Serie','Año','VidaUtilResidual','Clasificación','ENU/Baja','Observación','Frecuencia','Responsable'];
  for (let m = 0; m < 12; m++) headersPmp.push(MESES[m]);

  sPmp.getRange(7, 1, 1, headersPmp.length).setValues([headersPmp]).setFontWeight('bold');

  const headersReg = ['Familia','ID','N°Carpeta','N°Inventario','Equipo','Servicio','Unidad','Ubicación','Procedencia','Marca','Modelo','Serie','Año','VidaUtilResidual','Clasificación','ENU/Baja','Observación','Frecuencia','Responsable'];
  for (let m = 0; m < 12; m++) { headersReg.push(MESES[m] + ' P'); headersReg.push(MESES[m] + ' R'); }
  sReg.getRange(7, 1, 1, headersReg.length).setValues([headersReg]).setFontWeight('bold');

  const pmpRows = []; const regRows = [];
  for (let i = 1; i <= 20; i++) {
    const fam = familias[i % familias.length];
    const inv = 'INV' + String(10000 + i);
    const eq = fam;
    const serv = servicios[i % servicios.length];
    const ubic = ubicaciones[i % ubicaciones.length];
    const marca = marcas[i % marcas.length];
    const row = [familiaDe(fam), i, 'C-' + i, inv, eq, serv, '', ubic, 'Donación', marca, 'M-' + i, 'S-' + i, 2020 + (i % 5), 8, 'Crítico', '', '', i % 2 === 0 ? 'Anual' : 'Semestral', ''];
    // meses prog
    for (let m = 0; m < 12; m++) row.push(m === (i % 12) || m === ((i + 6) % 12) ? 'X' : '');
    pmpRows.push(row);

    const rrow = [familiaDe(fam), i, 'C-' + i, inv, eq, serv, '', ubic, 'Donación', marca, 'M-' + i, 'S-' + i, 2020 + (i % 5), 8, 'Crítico', '', '', i % 2 === 0 ? 'Anual' : 'Semestral', ''];
    for (let m = 0; m < 12; m++) {
      const programado = m === (i % 12) || m === ((i + 6) % 12);
      rrow.push(programado ? 'X' : '');
      if (programado && m < new Date().getMonth()) rrow.push(i % 5 === 0 ? 'C1' : 'Si');
      else rrow.push('');
    }
    regRows.push(rrow);
  }
  sPmp.getRange(8, 1, pmpRows.length, headersPmp.length).setValues(pmpRows);
  sReg.getRange(8, 1, regRows.length, headersReg.length).setValues(regRows);

  // Pendientes ficticios
  for (let i = 0; i < 10; i++) {
    const equipoKey = 'inv:INV' + String(10000 + i + 1);
    savePendiente({
      equipoKey: equipoKey,
      descripcion: 'Tarea de prueba #' + (i + 1),
      fechaCompromiso: Utilities.formatDate(new Date(Date.now() + (i - 3) * 86400000), 'America/Santiago', 'yyyy-MM-dd'),
      ejecutor: EJECUTORES_DEFAULT[i % EJECUTORES_DEFAULT.length],
      prioridad: ['alta','media','baja'][i % 3],
      etiquetas: ['prueba','sintetico'],
    });
  }

  invalidateCache();
  return { equipos: pmpRows.length, pendientes: 10 };
});

// ====================================================================
// onEdit trigger (opcional, configurable en Config)
// ====================================================================
function onEditTrigger(e) {
  try {
    if (!e || !e.range) return;
    const sheetName = e.range.getSheet().getName();
    if (sheetName !== SHEETS.REG) return;
    if (String(getConfigValue('ONEDIT_AUTO')) !== 'true') return;
    invalidateCache();
    // Marcar timestamp para que la app sepa que hay cambios
    setConfigValue('LAST_EDIT_REGISTRO', nowIso());
  } catch (err) {}
}

// ====================================================================
// Enviar recordatorios (opcional manual)
// ====================================================================
const enviarRecordatorios = wrap(function() {
  const s = ensureSheet(SHEETS.PENDIENTES, HEADERS.PENDIENTES);
  const rows = s.getDataRange().getValues();
  const hoy = todayIso();
  const byEjecutor = {};
  for (let i = 1; i < rows.length; i++) {
    const estado = rows[i][8]; const ejecutor = rows[i][5]; const fc = rows[i][4];
    if (estado === 'abierto' && fc && String(fc) < hoy && ejecutor) {
      if (!byEjecutor[ejecutor]) byEjecutor[ejecutor] = [];
      byEjecutor[ejecutor].push({ desc: rows[i][2], fc: fc });
    }
  }
  const me = userEmail();
  let n = 0;
  for (const ej in byEjecutor) {
    const lista = byEjecutor[ej].map(function(p) { return '- ' + p.desc + ' (compromiso: ' + p.fc + ')'; }).join('\n');
    MailApp.sendEmail({
      to: me,
      subject: 'MP 2026 — Pendientes vencidos de ' + ej,
      body: 'Pendientes vencidos asignados a ' + ej + ':\n\n' + lista,
    });
    n++;
  }
  return { ejecutoresNotificados: n };
});
