/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Setup.gs
 *  inicializarV2(): prepara el Maestro. Solo CREA o COMPLETA; nunca borra.
 * ═══════════════════════════════════════════════════════════════════════════
 *  Ejecutar una vez desde el editor de Apps Script (y de nuevo si se borra
 *  alguna hoja). Es seguro repetirla.
 */

function inicializarV2() {
  const ss = spreadsheet_();
  const reporte = [];

  // 1. Zona horaria de la hoja = zona del script (evita fechas a las 23:00 del día anterior).
  if (ss.getSpreadsheetTimeZone() !== TZ) {
    reporte.push('Zona horaria de la hoja: ' + ss.getSpreadsheetTimeZone() + ' → ' + TZ);
    ss.setSpreadsheetTimeZone(TZ);
    _ssTz = null;
  } else reporte.push('Zona horaria de la hoja: ok (' + TZ + ')');

  // 2. Hojas con estructura propia
  asegurarHoja_(ss, SHEETS.CACHE_REP1, 'CACHE — ÚLTIMO REP1 CARGADO (vencimientos de la tanda)', REP1_HEADERS, 4, reporte);
  asegurarHoja_(ss, SHEETS.CACHE_REP9, 'CACHE — ÚLTIMO REP9 CARGADO (saldos de cartera)', REP9_HEADERS, 4, reporte);
  asegurarHoja_(ss, SHEETS.BITACORA, 'BITÁCORA DE ENVÍOS DE AVISOS DE COBRANZA', BITACORA_HEADERS, 2, reporte);
  asegurarHoja_(ss, SHEETS.CORTES, null, ['Corte', 'Reporte', 'Fecha y hora de carga', 'Cargado por', 'Archivo', 'Filas', 'Vencimiento mín.', 'Vencimiento máx.', 'Rango declarado', 'Control (Importe / Saldo total)'], 1, reporte);
  asegurarHoja_(ss, SHEETS.AJUSTES, null, ['ID', 'Fecha', 'Usuario', 'Línea', 'Vencimiento (opcional)', 'Tipo', 'Monto', 'Motivo', 'Estado', 'Corte Rep9', 'Aprobó', 'Nota'], 1, reporte);
  asegurarHoja_(ss, SHEETS.INHABILES, null, ['Fecha', 'Descripción', 'Origen', 'Estado (ACTIVO / QUITAR)'], 1, reporte);
  asegurarHoja_(ss, SHEETS.USUARIOS, null, ['Correo', 'Nombre', 'Rol (COORDINADORA / GERENTE / AUDITORIA / ADMIN)', 'Activo (SI / NO)'], 1, reporte);
  asegurarHoja_(ss, SHEETS.CAMBIOS, null, ['Fecha', 'Usuario', 'Hoja', 'Línea', 'Campo', 'Antes', 'Después'], 1, reporte);

  // 3. Encabezados nuevos de la bitácora (columnas J a T) sin tocar las existentes
  const bit = ss.getSheetByName(SHEETS.BITACORA);
  const actuales = bit.getRange(2, 1, 1, BITACORA_COLS).getValues()[0];
  let agregados = 0;
  BITACORA_HEADERS.forEach((h, i) => { if (!String(actuales[i] || '').trim()) { bit.getRange(2, i + 1).setValue(h); agregados++; } });
  if (agregados) { formatoEncabezado_(bit, 2, BITACORA_COLS); reporte.push('Bitácora: se agregaron ' + agregados + ' encabezados nuevos.'); }

  // 4. Parámetros nuevos en Config (sin pisar los existentes)
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  if (!cfg) reporte.push('⚠️ Config: NO ENCONTRADA — restáurala manualmente.');
  else {
    const existentes = new Set(cfg.getLastRow() >= 3 ? cfg.getRange(3, 1, cfg.getLastRow() - 2, 1).getValues().map(r => String(r[0]).trim()) : []);
    const nuevos = CONFIG_DEFAULTS.filter(d => !existentes.has(d[0]));
    if (nuevos.length) {
      cfg.getRange(cfg.getLastRow() + 1, 1, nuevos.length, 3).setValues(nuevos);
      reporte.push('Config: se agregaron ' + nuevos.length + ' parámetros (' + nuevos.map(n => n[0]).join(', ') + ').');
    } else reporte.push('Config: ya tiene todos los parámetros v2.');
    // Descripciones de parámetros heredados de la v1 cuyo significado cambió o que ya no se usan
    const DESC = {
      MODO_PRUEBA: 'TRUE = todos los correos llegan al buzón de prueba (MODO_PRUEBA_DESTINO) con prefijo [PRUEBA] y no consumen avisos reales. Permite fecha simulada.',
      VENTANA_DIAS_AVISO: 'Ya no se usa en la v2: la ventana la calcula la tanda con días hábiles.',
      DIAS_TIPO_T_CERO: 'Ya no se usa en la v2: el día del vencimiento no sale aviso.',
      HORA_TRIGGER_DIARIO: 'Ya no se usa en la v2: las horas se definen en HORA_RECORDATORIO, HORA_ESCALACION y HORA_ENVIO_AUTOMATICO.'
    };
    const lr = cfg.getLastRow();
    if (lr >= 3) {
      const vals = cfg.getRange(3, 1, lr - 2, 3).getValues();
      let cambios = 0;
      vals.forEach((r, i) => { const k = String(r[0]).trim(); if (DESC[k] && r[2] !== DESC[k]) { cfg.getRange(3 + i, 3).setValue(DESC[k]); cambios++; } });
      if (cambios) reporte.push('Config: se actualizó la descripción de ' + cambios + ' parámetros heredados.');
    }
  }

  // 5. Calendario de inhábiles visible (2026 a 2028)
  const cal = ss.getSheetByName(SHEETS.INHABILES);
  if (cal.getLastRow() < 2) {
    const filas = [];
    [2026, 2027, 2028].forEach(y => inhabilesPorRegla_(y).forEach(x => filas.push([fechaDeKey_(x[0]), x[1], y === 2026 ? 'CNBV 2026 (verificado)' : 'Regla — validar con calendario CNBV', 'ACTIVO'])));
    cal.getRange(2, 1, filas.length, 4).setValues(filas);
    cal.getRange(2, 1, filas.length, 1).setNumberFormat('dd/mm/yyyy');
    reporte.push('Calendario_Inhabiles: ' + filas.length + ' fechas (2026 verificado; 2027-2028 por regla, hay que validarlos).');
  }

  invalidarConfig_(); invalidarCalendario_();
  Logger.log(reporte.join('\n'));
  return { ok: true, reporte: reporte };
}

function asegurarHoja_(ss, nombre, titulo, headers, filaEncabezado, reporte) {
  let sh = ss.getSheetByName(nombre);
  const creada = !sh;
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() >= filaEncabezado && !creada) { reporte.push('✓ ' + nombre + ': existe'); return sh; }
  if (titulo) {
    sh.getRange(1, 1).setValue(titulo);
    sh.getRange(1, 1, 1, headers.length).merge().setFontSize(14).setFontWeight('bold').setFontColor('#515151')
      .setBackground('#FDB913').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(1, 28);
  }
  sh.getRange(filaEncabezado, 1, 1, headers.length).setValues([headers]);
  formatoEncabezado_(sh, filaEncabezado, headers.length);
  sh.setFrozenRows(filaEncabezado);
  for (let c = 1; c <= headers.length; c++) sh.setColumnWidth(c, 120);
  reporte.push((creada ? '✅ ' : '🔧 ') + nombre + (creada ? ': creada' : ': estructura aplicada'));
  return sh;
}

function formatoEncabezado_(sh, fila, n) {
  sh.getRange(fila, 1, 1, n).setFontWeight('bold').setBackground('#515151').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setFontFamily('Arial').setFontSize(10);
  sh.setRowHeight(fila, 32);
}

/**
 * Compatibilidad: si quedó instalado el trigger de la v1 (cronEnvioDiario)
 * se reemplaza por el flujo seguro de la v2.
 */
function cronEnvioDiario() { envioAutomaticoDiario(); }
