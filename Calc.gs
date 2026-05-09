/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA — Motor de Cálculo (Calc.gs)
 * ═══════════════════════════════════════════════════════════════════════════
 *  Replica las fórmulas de la Hoja 1 vieja, con match exacto por línea
 *  y manejo explícito de datos faltantes.
 *
 *  ESTRUCTURA REP1 (Cache_Rep1, 10 cols):
 *    A: Fecha Vencimiento  B: Línea  C: Nombre  D: Capital  E: Interés
 *    F: Otros  G: IVA  H: Importe  I: Moneda
 *
 *  MODO PRUEBA (controlable desde Config):
 *    Cuando MODO_PRUEBA = TRUE: cualquier línea con venc ≥ mañana → T-1 enviable.
 *    Cuando MODO_PRUEBA = FALSE: solo días=5/1/0 son enviables.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─── CONSTANTES DEL DOMINIO ────────────────────────────────────────────────
const FACTOR_TASA_MORATORIA = 2;
const BASE_DIAS_ANIO = 360;
const VENTANA_DIAS_AVISO_DEFAULT = 5;

// Mapeo Días → Tipo de Aviso (modo producción)
const TIPO_AVISO = { 5: 'T-5', 1: 'T-1', 0: 'T+0' };

// ─── ÍNDICES DE COLUMNAS Cache_Rep1 (NUEVA estructura: 10 cols) ────────────
const REP1 = {
  FECHA_VENC: 0,  // A
  LINEA:      1,  // B
  NOMBRE:     2,  // C
  CAPITAL:    3,  // D
  INTERES:    4,  // E
  OTROS:      5,  // F
  IVA:        6,  // G
  IMPORTE:    7,  // H
  MONEDA:     8   // I
};

// ─── ÍNDICES DE COLUMNAS Cache_Rep9 ────────────────────────────────────────
const REP9 = {
  LINEA:           0,
  NOMBRE:          1,
  MONEDA:          3,
  CAP_VENCIDO:    10,
  INT_VENCIDO:    15,
  IVA_INT_VENCIDO: 16,
  MORATORIOS:     17,
  IVA_MORATORIOS: 18,
  MOR_CONT:       19,
  IVA_MOR_CONT:   20
};

// ─── API: ESTADO DEL MODO PRUEBA ──────────────────────────────────────────

function getModoPrueba() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  if (!sh) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 3) return false;

  const data = sh.getRange(3, 1, lastRow - 2, 2).getValues();
  for (const row of data) {
    if (String(row[0]).trim() === 'MODO_PRUEBA') {
      const v = row[1];
      if (typeof v === 'boolean') return v;
      const s = String(v).trim().toUpperCase();
      return s === 'TRUE' || s === 'VERDADERO' || s === '1' || s === 'SI' || s === 'SÍ';
    }
  }

  // No existe → crearla
  const newRow = lastRow + 1;
  sh.getRange(newRow, 1, 1, 3).setValues([[
    'MODO_PRUEBA', false,
    'Cuando es TRUE, cualquier vencimiento futuro (≥ mañana) se trata como T-1 enviable. Solo para pruebas.'
  ]]);
  sh.getRange(newRow, 2).setBackground('#FFF8E1');
  sh.getRange(newRow, 1).setFontWeight('bold');

  return false;
}

function setModoPrueba(activo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.CONFIG);
  if (!sh) return { ok: false, error: 'Hoja Config no encontrada.' };

  const lastRow = sh.getLastRow();
  const data = sh.getRange(3, 1, Math.max(1, lastRow - 2), 2).getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'MODO_PRUEBA') {
      sh.getRange(3 + i, 2).setValue(Boolean(activo));
      return { ok: true, modoPrueba: Boolean(activo) };
    }
  }

  getModoPrueba();
  return setModoPrueba(activo);
}

function getVentanaDiasAviso_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEETS.CONFIG);
    if (!sh) return VENTANA_DIAS_AVISO_DEFAULT;
    const lastRow = sh.getLastRow();
    if (lastRow < 3) return VENTANA_DIAS_AVISO_DEFAULT;
    const data = sh.getRange(3, 1, lastRow - 2, 2).getValues();
    for (const row of data) {
      if (String(row[0]).trim() === 'VENTANA_DIAS_AVISO') {
        const n = Number(row[1]);
        return (isNaN(n) || n < 1) ? VENTANA_DIAS_AVISO_DEFAULT : Math.round(n);
      }
    }
    // Crear si no existe
    const newRow = sh.getLastRow() + 1;
    sh.getRange(newRow, 1, 1, 3).setValues([[
      'VENTANA_DIAS_AVISO', VENTANA_DIAS_AVISO_DEFAULT,
      'Días hacia adelante a considerar para el cálculo (5 producción, 30+ pruebas).'
    ]]);
    sh.getRange(newRow, 2).setBackground('#FFF8E1');
    sh.getRange(newRow, 1).setFontWeight('bold');
    return VENTANA_DIAS_AVISO_DEFAULT;
  } catch (err) {
    return VENTANA_DIAS_AVISO_DEFAULT;
  }
}

// ─── API PRINCIPAL ─────────────────────────────────────────────────────────

function calcularAvisos() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const rep1Sheet = ss.getSheetByName(SHEETS.CACHE_REP1);
    const rep9Sheet = ss.getSheetByName(SHEETS.CACHE_REP9);
    if (!rep1Sheet || !rep9Sheet) {
      return { ok: false, error: 'Hojas de cache no encontradas.' };
    }
    if (rep1Sheet.getLastRow() < 5) {
      return { ok: false, error: 'No hay Rep1 cargado.' };
    }
    if (rep9Sheet.getLastRow() < 5) {
      return { ok: false, error: 'No hay Rep9 cargado.' };
    }

    const modoPrueba = getModoPrueba();
    const ventanaDias = modoPrueba
      ? Math.max(60, getVentanaDiasAviso_())
      : getVentanaDiasAviso_();

    // Leer datos
    const rep1Data = leerCache_(rep1Sheet, 9);   // 9 cols del nuevo Cache_Rep1
    const rep9Data = leerCache_(rep9Sheet, 30);
    const tasasMap = leerTasas_(ss);
    const correosMap = leerCorreos_(ss);

    const rep9Map = new Map();
    for (const r of rep9Data) {
      const linea = normLinea_(r[REP9.LINEA]);
      if (linea) rep9Map.set(linea, r);
    }

    const fechaReporte = new Date();
    fechaReporte.setHours(0, 0, 0, 0);
    const limiteFecha = new Date(fechaReporte);
    limiteFecha.setDate(limiteFecha.getDate() + ventanaDias);

    const avisos = [];
    const warnings = [];

    for (const r of rep1Data) {
      const linea = normLinea_(r[REP1.LINEA]);
      if (!linea) continue;

      const fechaVenc = parseFecha_(r[REP1.FECHA_VENC]);
      if (!fechaVenc) {
        warnings.push(`Línea ${linea}: fecha inválida.`);
        continue;
      }

      if (fechaVenc.getTime() > limiteFecha.getTime()) continue;

      const dias = diferenciaDias_(fechaReporte, fechaVenc);

      // ─── DETERMINAR TIPO DE AVISO ───
      let tipoAviso, elegibleAviso;
      if (modoPrueba) {
        if (dias >= 1) { tipoAviso = 'T-1'; elegibleAviso = true; }
        else if (dias === 0) { tipoAviso = 'T+0'; elegibleAviso = true; }
        else { tipoAviso = 'Vencido'; elegibleAviso = false; }
      } else {
        tipoAviso = TIPO_AVISO[dias] || (dias < 0 ? 'Vencido' : 'Fuera de rango');
        elegibleAviso = (dias === 5 || dias === 1 || dias === 0);
      }

      // ─── CÁLCULOS ───
      const capital   = num_(r[REP1.CAPITAL]);
      const intereses = num_(r[REP1.INTERES]);
      const otros     = num_(r[REP1.OTROS]);
      const iva       = num_(r[REP1.IVA]);
      const importe   = num_(r[REP1.IMPORTE]);
      const moneda    = String(r[REP1.MONEDA] || 'MXN').trim().toUpperCase();

      const r9 = rep9Map.get(linea);
      let capVencido = 0, intVencidos = 0, moratoriosAcum = 0;
      let sinRep9 = false;

      if (r9) {
        capVencido     = num_(r9[REP9.CAP_VENCIDO]);
        intVencidos    = num_(r9[REP9.INT_VENCIDO]) + num_(r9[REP9.IVA_INT_VENCIDO]);
        moratoriosAcum = num_(r9[REP9.MORATORIOS]) + num_(r9[REP9.IVA_MORATORIOS]) +
                         num_(r9[REP9.MOR_CONT]) + num_(r9[REP9.IVA_MOR_CONT]);
      } else {
        sinRep9 = true;
      }

      const tasaInfo = tasasMap.get(linea);
      const tasaContrato = tasaInfo ? tasaInfo.tasa : 0;
      const tasaMoratoria = tasaContrato * FACTOR_TASA_MORATORIA;

      const moratoriosPeriodo = capVencido > 0 && dias > 0
        ? capVencido * tasaMoratoria / BASE_DIAS_ANIO * dias
        : 0;

      const total = capital + intereses + otros + iva +
                    capVencido + intVencidos + moratoriosAcum + moratoriosPeriodo;

      const contacto = correosMap.get(linea) || {};

      avisos.push({
        linea: linea,
        nombre: r[REP1.NOMBRE] || (r9 ? r9[REP9.NOMBRE] : '') || '',
        fechaVenc: fechaVenc.toISOString(),
        dias: dias,
        tipoAviso: tipoAviso,
        elegibleAviso: elegibleAviso,
        sinRep9: sinRep9,
        sinTasa: !tasaInfo,
        sinCorreo: !contacto.correo,
        moneda: moneda,                  // ← NUEVO: MXN o USD
        capital: capital,
        intereses: intereses,
        otros: otros,
        iva: iva,
        importe: importe,
        capVencido: capVencido,
        intVencidos: intVencidos,
        moratoriosAcum: moratoriosAcum,
        tasaContrato: tasaContrato,
        tasaMoratoria: tasaMoratoria,
        moratoriosPeriodo: round2_(moratoriosPeriodo),
        total: round2_(total),
        correo: contacto.correo || '',
        cuentaSTP: contacto.cuentaSTP || '',
        cliente: contacto.cliente || ''
      });

      if (sinRep9) warnings.push(`Línea ${linea}: sin saldo en Rep9.`);
    }

    avisos.sort((a, b) => {
      const da = new Date(a.fechaVenc).getTime();
      const db = new Date(b.fechaVenc).getTime();
      if (da !== db) return da - db;
      return String(a.linea).localeCompare(String(b.linea));
    });

    const elegibles = avisos.filter(a => a.elegibleAviso);
    const stats = {
      total: avisos.length,
      elegibles: elegibles.length,
      porTipo: {
        'T-5': elegibles.filter(a => a.tipoAviso === 'T-5').length,
        'T-1': elegibles.filter(a => a.tipoAviso === 'T-1').length,
        'T+0': elegibles.filter(a => a.tipoAviso === 'T+0').length
      },
      porMoneda: {
        'MXN': elegibles.filter(a => a.moneda === 'MXN').length,
        'USD': elegibles.filter(a => a.moneda === 'USD').length
      },
      sumaTotalMXN: round2_(elegibles.filter(a => a.moneda === 'MXN').reduce((s, a) => s + a.total, 0)),
      sumaTotalUSD: round2_(elegibles.filter(a => a.moneda === 'USD').reduce((s, a) => s + a.total, 0)),
      sumaCapVencido: round2_(elegibles.reduce((s, a) => s + a.capVencido, 0)),
      sumaMoratorios: round2_(elegibles.reduce((s, a) => s + a.moratoriosAcum + a.moratoriosPeriodo, 0)),
      sinRep9: elegibles.filter(a => a.sinRep9).length,
      sinCorreo: elegibles.filter(a => a.sinCorreo).length,
      sinTasa: elegibles.filter(a => a.sinTasa).length
    };

    return {
      ok: true,
      modoPrueba: modoPrueba,
      ventanaDias: ventanaDias,
      fechaReporte: fechaReporte.toISOString(),
      avisos: avisos,
      stats: stats,
      warnings: warnings.slice(0, 50)
    };

  } catch (err) {
    return { ok: false, error: err.message, stack: err.stack };
  }
}

// ─── HELPERS DE LECTURA ────────────────────────────────────────────────────

function leerCache_(sheet, numCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return [];
  const data = sheet.getRange(5, 1, lastRow - 4, numCols).getValues();
  return data.filter(r => r.some(c => c !== null && c !== undefined && c !== ''));
}

function leerTasas_(ss) {
  const sh = ss.getSheetByName(SHEETS.TASAS);
  if (!sh) return new Map();
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return new Map();
  const data = sh.getRange(3, 1, lastRow - 2, 3).getValues();
  const map = new Map();
  for (const r of data) {
    const linea = normLinea_(r[0]);
    if (!linea) continue;
    map.set(linea, { nombre: r[1], tasa: num_(r[2]) });
  }
  return map;
}

function leerCorreos_(ss) {
  const sh = ss.getSheetByName(SHEETS.CORREOS);
  if (!sh) return new Map();
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return new Map();
  const data = sh.getRange(3, 1, lastRow - 2, 5).getValues();
  const map = new Map();
  for (const r of data) {
    const linea = normLinea_(r[1]);
    if (!linea) continue;
    map.set(linea, {
      cliente: r[2],
      correo: String(r[3] || '').trim(),
      cuentaSTP: r[4]
    });
  }
  return map;
}

// ─── HELPERS DE NORMALIZACIÓN ──────────────────────────────────────────────

function normLinea_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return String(Math.round(v));
  const s = String(v).trim();
  if (s === '') return null;
  const n = Number(s);
  if (!isNaN(n) && isFinite(n)) return String(Math.round(n));
  return s;
}

function num_(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function parseFecha_(v) {
  if (!v) return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  if (typeof v === 'string') {
    const d = new Date(v);
    if (!isNaN(d.getTime())) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
  }
  return null;
}

function diferenciaDias_(desde, hasta) {
  const ms = hasta.getTime() - desde.getTime();
  return Math.round(ms / (1000 * 60 * 60 * 24));
}

function round2_(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}
