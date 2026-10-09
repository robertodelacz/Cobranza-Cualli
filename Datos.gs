/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Datos.gs
 *  Reportes (Rep1/Rep9), cortes, catálogos, ajustes y bitácora.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─── ESTRUCTURAS DE LOS REPORTES ───────────────────────────────────────────

const REP1_HEADERS = [
  'Fecha Vencimiento', 'Línea de Crédito', 'Nombre de Cliente', 'Capital',
  'Interés', 'Otros', 'IVA', 'Importe', 'Moneda'
];

const REP9_HEADERS = [
  'Línea de crédito', 'Nombre Cliente', 'Num. Cliente_ID', 'Moneda',
  'Fecha Alta', 'Fijo hasta', '*Saldo disp. de la línea', '*Monto línea',
  'Monto del crédito', 'Cap. Vigente', 'Cap. Vencido', '*Saldo Cap.',
  'IVA Cliente', 'Int. Vigente', 'IVA Int. Vigente', 'Int. Vencido',
  'IVA Int. Vencido', 'Moratorios', 'IVA Moratorios', 'Mor. Contabilizados',
  'IVA Mor. Contabilizados', '*Saldo Int.', 'IVA Comisiones',
  '*Saldo Comisiones', '*IVA Saldo Comisiones', '*Saldo Vigente',
  '*Saldo Vencido', '*SALDO TOTAL', 'Pagos No Aplicados', 'Promotor'
];

const REP1 = { FECHA_VENC: 0, LINEA: 1, NOMBRE: 2, CAPITAL: 3, INTERES: 4, OTROS: 5, IVA: 6, IMPORTE: 7, MONEDA: 8 };
const REP9 = {
  LINEA: 0, NOMBRE: 1, CLIENTE_ID: 2, MONEDA: 3, CAP_VIGENTE: 9, CAP_VENCIDO: 10, SALDO_CAP: 11, INT_VENCIDO: 15, IVA_INT_VENCIDO: 16,
  MORATORIOS: 17, IVA_MORATORIOS: 18, MOR_CONT: 19, IVA_MOR_CONT: 20, SALDO_INT: 21, SALDO_VIGENTE: 25, SALDO_VENCIDO: 26, SALDO_TOTAL: 27,
  PAGOS_NO_APLICADOS: 28
};

// ─── NORMALIZACIÓN DE NÚMEROS Y LÍNEAS ─────────────────────────────────────

function num_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v).trim();
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
  s = s.replace(/[$\s,]/g, '');
  const n = Number(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}
function round2_(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

function normLinea_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return String(Math.round(v));
  const s = String(v).trim();
  if (s === '') return null;
  const n = Number(s);
  if (!isNaN(n) && isFinite(n)) return String(Math.round(n));
  return s;
}

function colToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/** Minúsculas, sin acentos y con espacios colapsados: para comparar encabezados. */
function normEnc_(s) { return String(s === null || s === undefined ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim(); }

function esCorreo_(s) { return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(s).trim()); }
function parsearDestinatarios_(raw) {
  if (!raw) return [];
  return String(raw).split(/[,;]/).map(s => s.trim()).filter(s => s.length > 0 && esCorreo_(s));
}

// ─── REP1: DETECCIÓN DE ENCABEZADO Y NORMALIZACIÓN ─────────────────────────

/**
 * Toma el Rep1 tal cual lo leyó el navegador y devuelve { allRows, dataRows }
 * con las 9 columnas estándar. Busca el encabezado en las primeras 10 filas y
 * descarta la segunda columna "Linea de Crédito" que el reporte trae repetida.
 */
function normalizarRep1Raw_(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return { ok: false, error: 'Rep1: el archivo no contiene datos suficientes.' };
  }
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const fila = rows[i].map(c => normEnc_(c));
    if (fila.some(c => c.indexOf('fecha vencimiento') >= 0 || c === 'fecha venc.')) { headerRowIdx = i; break; }
  }
  if (headerRowIdx === -1) {
    return { ok: false, error: 'Rep1: no se encontró el encabezado "Fecha Vencimiento" en las primeras 10 filas.' };
  }
  const rawHeaders = rows[headerRowIdx].map(c => String(c || '').trim());
  const findIdx = (target) => {
    const t = normEnc_(target);
    for (let i = 0; i < rawHeaders.length; i++) {
      if (normEnc_(rawHeaders[i]) === t) return i;
    }
    return -1;
  };
  const idx = {
    fecha: findIdx('fecha vencimiento'), linea: findIdx('linea de crédito'), nombre: findIdx('nombre de cliente'),
    capital: findIdx('capital'), interes: findIdx('interés'), otros: findIdx('otros'),
    iva: findIdx('iva'), importe: findIdx('importe'), moneda: findIdx('moneda')
  };
  const nombres = { fecha: 'Fecha Vencimiento', linea: 'Linea de Crédito', nombre: 'Nombre de Cliente', capital: 'Capital',
    interes: 'Interés', otros: 'Otros', iva: 'IVA', importe: 'Importe', moneda: 'Moneda' };
  const faltantes = Object.keys(idx).filter(k => idx[k] < 0).map(k => nombres[k]);
  if (faltantes.length) {
    return { ok: false, error: 'Rep1: faltan columnas requeridas: ' + faltantes.join(', '),
             mismatches: faltantes.map(n => 'Columna no encontrada: "' + n + '"') };
  }
  const dataRows = [];
  rows.slice(headerRowIdx + 1).forEach(row => {
    if (!row || !row.some(c => c !== null && c !== undefined && c !== '')) return;
    const fechaCell = row[idx.fecha], lineaCell = row[idx.linea];
    if (!fechaCell || !lineaCell) return;
    if (typeof lineaCell === 'string' && !/^\d+$/.test(String(lineaCell).trim())) return;   // subtotales
    dataRows.push([
      fechaCell, lineaCell, row[idx.nombre] || '',
      num_(row[idx.capital]), num_(row[idx.interes]), num_(row[idx.otros]), num_(row[idx.iva]), num_(row[idx.importe]),
      String(row[idx.moneda] || 'MXN').trim().toUpperCase()
    ]);
  });
  if (!dataRows.length) return { ok: false, error: 'Rep1: no se encontraron filas de datos válidas después del encabezado.' };
  return { ok: true, allRows: [REP1_HEADERS].concat(dataRows), dataRows: dataRows };
}

/** Valida el Rep9 (30 columnas con encabezado exacto en la fila 1). */
function validarRep9Raw_(rows) {
  if (!Array.isArray(rows) || rows.length < 2) return { ok: false, error: 'Rep9: el archivo no contiene datos.' };
  const headers = rows[0].map(h => String(h || '').trim());
  if (headers.length < REP9_HEADERS.length) {
    return { ok: false, error: 'Rep9: se esperaban ' + REP9_HEADERS.length + ' columnas, se recibieron ' + headers.length + '.' };
  }
  const mismatches = [];
  REP9_HEADERS.forEach((h, i) => {
    const a = normEnc_(headers[i]), b = normEnc_(h);
    if (a !== b) mismatches.push('Col ' + (i + 1) + ': esperaba "' + h + '", recibió "' + headers[i] + '"');
  });
  if (mismatches.length) return { ok: false, error: 'Rep9: las columnas no coinciden con el formato esperado.', mismatches: mismatches };
  const dataRows = rows.slice(1).filter(r => r.some(c => c !== null && c !== undefined && c !== ''));
  return { ok: true, allRows: rows.slice(0, 1).concat(dataRows), dataRows: dataRows };
}

// ─── CORTES (cada carga queda registrada) ──────────────────────────────────

/** Lee los metadatos que se escriben en la fila 2 del caché. */
function leerCorteMeta_(sh) {
  const v = sh.getRange(2, 1, 1, 10).getValues()[0];
  let ts = v[1];
  if (Object.prototype.toString.call(ts) === '[object Date]') ts = Utilities.formatDate(ts, ssTz_(), 'yyyy-MM-dd HH:mm:ss');
  ts = String(ts || '').replace(/^'/, '').trim();
  const fechaKey = /^\d{4}-\d{2}-\d{2}/.test(ts) ? ts.slice(0, 10) : null;
  return {
    cargado: !!fechaKey, fechaHora: ts, fechaKey: fechaKey, usuario: String(v[3] || ''),
    corteId: String(v[5] || ''), archivo: String(v[7] || ''), rango: String(v[9] || ''),
    filas: Math.max(0, sh.getLastRow() - 4)
  };
}

function registrarCorte_(tipo, corteId, ahora, usuario, archivo, filas, minK, maxK, rango, controlTotal) {
  const sh = spreadsheet_().getSheetByName(SHEETS.CORTES);
  if (!sh) return;
  const r = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(r, 1, 1, 10).setValues([[corteId, tipo, ahora, usuario, archivo, filas, minK || '', maxK || '', rango || '', controlTotal]]);
  sh.getRange(r, 3).setNumberFormat('@');
}

/**
 * Guarda el reporte validado en su caché (reemplaza el anterior) y deja
 * constancia en la hoja Cortes. Devuelve el resumen del corte.
 */
function guardarCache_(tipo, rows, meta) {
  const esRep1 = tipo === 'rep1';
  const nombreHoja = esRep1 ? SHEETS.CACHE_REP1 : SHEETS.CACHE_REP9;
  const headers = esRep1 ? REP1_HEADERS : REP9_HEADERS;
  const sh = spreadsheet_().getSheetByName(nombreHoja);
  if (!sh) return { ok: false, error: 'Hoja "' + nombreHoja + '" no encontrada. Ejecuta inicializarV2().' };

  const usuario = usuarioActual_().email;
  const ahora = ahoraMX_();
  const corteId = (esRep1 ? 'R1-' : 'R9-') + ahora.replace(/[-: ]/g, '').slice(0, 14) + '-' + Math.floor(Math.random() * 1296).toString(36).padStart(2, '0');

  const dataRows = rows.slice(1).filter(r => r.some(c => c !== null && c !== undefined && c !== ''));
  if (!dataRows.length) return { ok: false, error: tipo.toUpperCase() + ': no hay filas con datos.' };

  const n = headers.length;
  let minK = null, maxK = null, control = 0;
  const salida = dataRows.map(r => {
    const row = r.slice(0, n);
    while (row.length < n) row.push('');
    return row.map((cell, c) => {
      if (typeof cell === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(cell)) {
        const k = cell.slice(0, 10);
        if ((esRep1 && c === 0)) { if (!minK || k < minK) minK = k; if (!maxK || k > maxK) maxK = k; }
        return fechaDeKey_(k);
      }
      if (!esRep1 && c >= 6 && c <= 28) return cell === '' ? '' : num_(cell);   // montos del Rep9 como número
      return cell;
    });
  });
  salida.forEach(r => { control += esRep1 ? num_(r[REP1.IMPORTE]) : num_(r[REP9.SALDO_TOTAL]); });

  const last = sh.getLastRow();
  if (last > 4) sh.getRange(5, 1, last - 4, sh.getMaxColumns()).clearContent();
  sh.getRange(4, 1, 1, n).setValues([headers]);
  sh.getRange(5, 1, salida.length, n).setValues(salida);
  if (esRep1) sh.getRange(5, 1, salida.length, 1).setNumberFormat('dd/MM/yyyy');
  else sh.getRange(5, 5, salida.length, 2).setNumberFormat('dd/MM/yyyy');

  const rango = (meta && meta.rangoDesde && meta.rangoHasta) ? (meta.rangoDesde + ' a ' + meta.rangoHasta) : '';
  sh.getRange(2, 1, 1, 10).setValues([['Fecha de carga:', "'" + ahora, 'Cargado por:', usuario, 'Corte:', corteId,
                                       'Archivo:', (meta && meta.archivo) || '', 'Rango declarado:', rango]]);
  sh.getRange(2, 1, 1, 10).setFontFamily('Arial').setFontSize(10);
  [1, 3, 5, 7, 9].forEach(c => sh.getRange(2, c).setFontWeight('bold'));

  registrarCorte_(esRep1 ? 'Rep1' : 'Rep9', corteId, ahora, usuario, (meta && meta.archivo) || '', salida.length, minK, maxK, rango, round2_(control));
  return { ok: true, tipo: tipo, corteId: corteId, filas: salida.length, fechaHora: ahora, usuario: usuario,
           fechaMin: minK, fechaMax: maxK, controlTotal: round2_(control) };
}

function leerCache_(sh, numCols) {
  const lastRow = sh.getLastRow();
  if (lastRow < 5) return [];
  return sh.getRange(5, 1, lastRow - 4, numCols).getValues()
    .filter(r => r.some(c => c !== null && c !== undefined && c !== ''));
}

// ─── CATÁLOGOS ─────────────────────────────────────────────────────────────

function leerTasas_() {
  const sh = spreadsheet_().getSheetByName(SHEETS.TASAS);
  const map = new Map();
  if (!sh || sh.getLastRow() < 3) return map;
  sh.getRange(3, 1, sh.getLastRow() - 2, 3).getValues().forEach((r, i) => {
    const linea = normLinea_(r[0]);
    if (!linea) return;
    if (map.has(linea)) { map.get(linea).duplicada = true; return; }
    map.set(linea, { nombre: r[1], tasa: num_(r[2]), fila: i + 3, duplicada: false });
  });
  return map;
}

function leerContactos_() {
  const sh = spreadsheet_().getSheetByName(SHEETS.CORREOS);
  const map = new Map();
  if (!sh || sh.getLastRow() < 3) return map;
  sh.getRange(3, 1, sh.getLastRow() - 2, 5).getValues().forEach((r, i) => {
    const linea = normLinea_(r[1]);
    if (!linea) return;
    if (map.has(linea)) { map.get(linea).duplicado = true; return; }
    map.set(linea, { numCliente: r[0], cliente: r[2], correo: String(r[3] || '').trim(), stp: String(r[4] || '').trim(), fila: i + 3, duplicado: false });
  });
  return map;
}

// ─── AJUSTES DE SALDO ──────────────────────────────────────────────────────
// Hoja Ajustes: A ID | B Fecha | C Usuario | D Línea | E Vencimiento (opcional) | F Tipo |
//               G Monto (con signo) | H Motivo | I Estado | J Corte Rep9 | K Aprobó | L Nota
// Un ajuste sólo aplica mientras el Rep9 cargado sea el mismo corte en que se capturó:
// al cargar un Rep9 nuevo se asume que ya trae el pago o la disposición.

function leerAjustes_() {
  const sh = spreadsheet_().getSheetByName(SHEETS.AJUSTES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues().map((r, i) => ({
    fila: i + 2, id: String(r[0] || ''), fecha: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd HH:mm') : String(r[1] || ''),
    usuario: String(r[2] || ''), linea: normLinea_(r[3]), vencimiento: fechaKeyDeHoja_(r[4]) || '',
    tipo: String(r[5] || ''), monto: num_(r[6]), motivo: String(r[7] || ''), estado: String(r[8] || ''),
    corte9: String(r[9] || ''), aprobo: String(r[10] || ''), nota: String(r[11] || '')
  })).filter(a => a.id);
}

// ─── BITÁCORA ──────────────────────────────────────────────────────────────
// A Timestamp | B Venc. | C Tipo | D Línea | E Cliente | F Correos | G Total | H Status | I Mensaje |
// J Usuario | K Corte Rep1 | L Corte Rep9 | M Moneda | N Días háb. | O Cuota | P Vencido | Q Moratorios | R Origen | S ID aviso | T Pago efectivo
// (B es la fecha NOMINAL del Rep1, que junto con la línea identifica la cuota.)

const BITACORA_COLS = 20;
const BITACORA_HEADERS = ['Timestamp Envío', 'Fecha Vencimiento', 'Tipo Aviso', 'Línea Crédito', 'Cliente', 'Correos Destino',
  'Total Aviso', 'Status', 'Mensaje / Error', 'Usuario', 'Corte Rep1', 'Corte Rep9', 'Moneda', 'Días háb. al pago',
  'Cuota', 'Vencido', 'Moratorios', 'Origen', 'ID aviso', 'Pago efectivo'];

function esTipoPrevio_(t) { return t === 'PREVENTIVO' || t === 'T-5'; }
function esTipoVispera_(t) { return t === 'VISPERA' || t === 'T-1' || t === 'T+0'; }

/**
 * Mapa linea|vencimiento → { previo, vispera, ultimo, reenvios }.
 * Sólo cuentan los ENVIADO reales: ENVIADO_PRUEBA, ERROR y OMITIDO no consumen el aviso.
 */
function leerBitacoraMapa_() {
  const map = new Map();
  const sh = spreadsheet_().getSheetByName(SHEETS.BITACORA);
  if (!sh || sh.getLastRow() < 3) return map;
  const total = sh.getLastRow() - 2;
  const filas = Math.min(total, 6000);
  const ini = sh.getLastRow() - filas + 1;
  sh.getRange(ini, 1, filas, 9).getValues().forEach(r => {
    const status = String(r[7] || '');
    if (status !== 'ENVIADO') return;
    const linea = normLinea_(r[3]);
    const venc = fechaKeyDeHoja_(r[1]);
    if (!linea || !venc) return;
    const ts = r[0];
    const tsKey = (Object.prototype.toString.call(ts) === '[object Date]') ? Utilities.formatDate(ts, TZ, 'yyyy-MM-dd') : fechaKeyDeHoja_(ts);
    if (!tsKey) return;
    const key = linea + '|' + venc;
    let e = map.get(key);
    if (!e) { e = { previo: null, vispera: null, ultimo: null, reenvios: 0 }; map.set(key, e); }
    const tipo = String(r[2] || '');
    const reg = { fechaKey: tsKey, total: num_(r[6]), tipo: tipo };
    if (esTipoPrevio_(tipo)) { if (!e.previo || tsKey < e.previo.fechaKey) e.previo = reg; }
    else if (esTipoVispera_(tipo)) { if (!e.vispera || tsKey < e.vispera.fechaKey) e.vispera = reg; }
    else if (tipo === 'REENVIO') e.reenvios++;
    if (!e.ultimo || tsKey >= e.ultimo.fechaKey) e.ultimo = reg;
  });
  return map;
}

function agregarBitacora_(filas) {
  if (!filas.length) return;
  const sh = spreadsheet_().getSheetByName(SHEETS.BITACORA);
  if (!sh) throw new Error('Hoja Bitacora_Envios no encontrada.');
  const r = Math.max(sh.getLastRow(), 2) + 1;
  sh.getRange(r, 1, filas.length, BITACORA_COLS).setValues(filas);
  sh.getRange(r, 1, filas.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange(r, 2, filas.length, 1).setNumberFormat('dd/mm/yyyy');
  sh.getRange(r, 7, filas.length, 1).setNumberFormat('$#,##0.00');
  sh.getRange(r, 15, filas.length, 3).setNumberFormat('$#,##0.00');
  sh.getRange(r, 20, filas.length, 1).setNumberFormat('dd/mm/yyyy');
}
