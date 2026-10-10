/**
 * COBRANZA PREVENTIVA v2 · Cartera
 * Vista por cliente: saldos del Rep9, próximas cuotas, avisos enviados y una
 * calculadora de "cuánto debe a una fecha" con la misma fórmula de los avisos.
 */

/** Últimos N movimientos de la bitácora, del más viejo al más nuevo. */
function bitacoraRegistros_(limite) {
  const sh = spreadsheet_().getSheetByName(SHEETS.BITACORA);
  if (!sh || sh.getLastRow() < 3) return [];
  const filas = Math.min(limite, sh.getLastRow() - 2);
  const ini = sh.getLastRow() - filas + 1;
  return sh.getRange(ini, 1, filas, BITACORA_COLS).getValues().map(r => ({
    timestamp: Object.prototype.toString.call(r[0]) === '[object Date]' ? Utilities.formatDate(r[0], TZ, "yyyy-MM-dd'T'HH:mm:ss") : '',
    fechaVenc: fechaKeyDeHoja_(r[1]) || '', tipo: String(r[2] || ''), linea: normLinea_(r[3]), cliente: String(r[4] || ''),
    correos: String(r[5] || ''), total: num_(r[6]), status: String(r[7] || ''), mensaje: String(r[8] || ''),
    usuario: String(r[9] || ''), corteRep1: String(r[10] || ''), corteRep9: String(r[11] || ''), moneda: String(r[12] || 'MXN'),
    diasHab: r[13], origen: String(r[17] || ''), pagoEfectivo: fechaKeyDeHoja_(r[19]) || ''
  }));
}

/** Saldos de una fila del Rep9, ya sumados como los usan los avisos. */
function saldosDeRep9_(x) {
  return {
    capVigente: round2_(num_(x[REP9.CAP_VIGENTE])),
    capVencido: round2_(num_(x[REP9.CAP_VENCIDO])),
    intVencidos: round2_(num_(x[REP9.INT_VENCIDO]) + num_(x[REP9.IVA_INT_VENCIDO])),
    moratorios: round2_(num_(x[REP9.MORATORIOS]) + num_(x[REP9.IVA_MORATORIOS]) + num_(x[REP9.MOR_CONT]) + num_(x[REP9.IVA_MOR_CONT])),
    saldoVigente: round2_(num_(x[REP9.SALDO_VIGENTE])),
    saldoVencido: round2_(num_(x[REP9.SALDO_VENCIDO])),
    saldoTotal: round2_(num_(x[REP9.SALDO_TOTAL])),
    pagosNoAplicados: round2_(num_(x[REP9.PAGOS_NO_APLICADOS]))
  };
}

function rep9Cargado_() {
  const sh9 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP9);
  if (!sh9 || sh9.getLastRow() < 5) return null;
  const rows = leerCache_(sh9, 30);
  const mapa = new Map();
  rows.forEach(r => { const l = normLinea_(r[REP9.LINEA]); if (l && !mapa.has(l)) mapa.set(l, r); });
  return { corte: leerCorteMeta_(sh9), rows: rows, mapa: mapa };
}

/** Todas las líneas del Rep9 con lo que necesita la vista de Cartera. */
function armarCartera_() {
  const r9 = rep9Cargado_();
  if (!r9) return { ok: false, codigo: 'SIN_REP9', error: 'Todavía no hay un Rep9 cargado. Súbelo en Hoy para ver la cartera.' };
  const contactos = leerContactos_(), tasas = leerTasas_();
  let cola = null;
  try { const c = calcularCola_(); if (c.ok) cola = c; } catch (e) {}
  const porLinea = new Map();
  if (cola) cola.items.forEach(i => { if (!porLinea.has(i.linea)) porLinea.set(i.linea, []); porLinea.get(i.linea).push(i); });
  const lineas = [];
  const totales = {};
  r9.mapa.forEach((x, linea) => {
    const s = saldosDeRep9_(x);
    const moneda = String(x[REP9.MONEDA] || 'MXN').trim().toUpperCase() || 'MXN';
    const cuotas = (porLinea.get(linea) || []).filter(i => i.diasHabiles >= 0).sort((a, b) => a.fechaPago < b.fechaPago ? -1 : 1);
    const prox = cuotas[0] || null;
    const contacto = contactos.get(linea) || null;
    const tieneCorreo = !!(contacto && parsearDestinatarios_(contacto.correo).length);
    lineas.push({
      linea: linea, cliente: (contacto && contacto.cliente) || String(x[REP9.NOMBRE] || ''), moneda: moneda,
      saldoTotal: s.saldoTotal, saldoVencido: s.saldoVencido, capVencido: s.capVencido,
      proxima: prox ? { fechaPago: prox.fechaPago, total: prox.total, estado: prox.estado, accion: prox.accion, key: prox.key } : null,
      enRep1: (porLinea.get(linea) || []).length > 0,
      tieneCorreo: tieneCorreo, tieneStp: !!(contacto && contacto.stp), tieneTasa: tasas.has(linea)
    });
    const t = totales[moneda] || (totales[moneda] = { saldoTotal: 0, saldoVencido: 0, lineas: 0, conVencido: 0 });
    t.saldoTotal += s.saldoTotal; t.saldoVencido += s.saldoVencido; t.lineas++;
    if (s.capVencido > 0.005 || s.saldoVencido > 0.005) t.conVencido++;
  });
  Object.keys(totales).forEach(m => { totales[m].saldoTotal = round2_(totales[m].saldoTotal); totales[m].saldoVencido = round2_(totales[m].saldoVencido); });
  lineas.sort((a, b) => b.saldoVencido - a.saldoVencido || (a.cliente < b.cliente ? -1 : 1));
  return { ok: true, corte: r9.corte, lineas: lineas, totales: totales, conAvisos: !!cola };
}

/** Todo lo que se sabe de una línea, para la ficha del cliente. */
function armarFichaCliente_(lineaIn) {
  const linea = normLinea_(lineaIn);
  if (!linea) return { ok: false, error: 'Indica la línea de crédito.' };
  const r9 = rep9Cargado_();
  const x = r9 ? r9.mapa.get(linea) : null;
  const contactos = leerContactos_(), tasas = leerTasas_();
  const contacto = contactos.get(linea) || null, tasa = tasas.get(linea) || null;
  let cola = null;
  try { const c = calcularCola_(); if (c.ok) cola = c; } catch (e) {}
  const cuotas = cola ? cola.items.filter(i => i.linea === linea).sort((a, b) => a.fechaPago < b.fechaPago ? -1 : 1) : [];
  const avisos = bitacoraRegistros_(2000).filter(b => b.linea === linea).reverse().slice(0, 40);
  const c9id = r9 ? r9.corte.corteId : '';
  const ajustes = leerAjustes_().filter(a => a.linea === linea).map(a => ({ id: a.id, fecha: a.fecha, tipo: a.tipo, monto: a.monto, motivo: a.motivo, estado: a.estado, vigente: a.corte9 === c9id && (a.estado === 'ACTIVO' || a.estado === 'PENDIENTE') }));
  if (!x && !contacto && !tasa && !cuotas.length && !avisos.length) return { ok: false, error: 'No encontré la línea ' + linea + ' en los reportes ni en los catálogos.' };
  return {
    ok: true, linea: linea, enRep9: !!x,
    cliente: (contacto && contacto.cliente) || (tasa && tasa.nombre) || (x ? String(x[REP9.NOMBRE] || '') : (cuotas[0] ? cuotas[0].cliente : '')),
    clienteId: x ? String(x[REP9.CLIENTE_ID] || '') : (contacto ? String(contacto.numCliente || '') : ''),
    moneda: x ? (String(x[REP9.MONEDA] || 'MXN').trim().toUpperCase() || 'MXN') : (cuotas[0] ? cuotas[0].moneda : 'MXN'),
    saldos: x ? saldosDeRep9_(x) : null, corte9: r9 ? r9.corte : null,
    cuotas: cuotas.map(i => ({ key: i.key, fechaNominal: i.fechaNominal, fechaPago: i.fechaPago, diasHabiles: i.diasHabiles, estado: i.estado, accion: i.accion, total: i.total, moneda: i.moneda, previo: i.previo, vispera: i.vispera })),
    avisos: avisos, ajustes: ajustes,
    contacto: { correo: contacto ? contacto.correo : '', stp: contacto ? contacto.stp : '' },
    tasa: tasa ? tasa.tasa : null
  };
}

/**
 * ¿Cuánto debe una línea el día X? Misma fórmula del aviso: cuotas del Rep1 que
 * vencen hasta X + saldo vencido del Rep9 + moratorios proyectados desde el corte
 * del Rep9 hasta X + ajustes vigentes. Es una estimación, no un estado de cuenta.
 */
function calcularSaldoAFecha_(lineaIn, fechaKey) {
  const linea = normLinea_(lineaIn);
  if (!linea) return { ok: false, error: 'Indica la línea de crédito.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fechaKey || ''))) return { ok: false, error: 'Elige una fecha válida.' };
  const r9 = rep9Cargado_();
  if (!r9) return { ok: false, error: 'Necesitas un Rep9 cargado para calcular.' };
  const x = r9.mapa.get(linea);
  if (!x) return { ok: false, error: 'La línea ' + linea + ' no aparece en el Rep9 cargado.' };
  const corteFecha = r9.corte.fechaKey;
  if (corteFecha && fechaKey < corteFecha) return { ok: false, error: 'La fecha no puede ser anterior al corte del Rep9 (' + fechaDMA_(corteFecha) + ').' };
  const s = saldosDeRep9_(x);
  const tasaInfo = leerTasas_().get(linea) || null;
  const factor = cfgNum_('FACTOR_TASA_MORATORIA', 2);
  const tasaMoratoria = tasaInfo ? tasaInfo.tasa * factor : 0;
  const dias = corteFecha ? Math.max(0, diffDays_(corteFecha, fechaKey)) : 0;
  const moratoriosProy = (s.capVencido > 0 && tasaMoratoria > 0 && dias > 0) ? round2_(s.capVencido * tasaMoratoria / BASE_DIAS_ANIO * dias) : 0;
  const sh1 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP1);
  const cuotas = [];
  if (sh1 && sh1.getLastRow() >= 5) {
    leerCache_(sh1, 9).forEach(r => {
      if (normLinea_(r[REP1.LINEA]) !== linea) return;
      const nominal = fechaKeyDeHoja_(r[REP1.FECHA_VENC]);
      if (!nominal || nominal > fechaKey) return;
      const monto = round2_(num_(r[REP1.CAPITAL]) + num_(r[REP1.INTERES]) + num_(r[REP1.OTROS]) + num_(r[REP1.IVA]));
      cuotas.push({ fecha: nominal, monto: monto });
    });
  }
  cuotas.sort((a, b) => a.fecha < b.fecha ? -1 : 1);
  const cuotasTotal = round2_(cuotas.reduce((t, c) => t + c.monto, 0));
  let ajustes = 0;
  leerAjustes_().forEach(a => { if (a.linea === linea && a.estado === 'ACTIVO' && a.corte9 === r9.corte.corteId) ajustes += a.monto; });
  ajustes = round2_(ajustes);
  const total = round2_(cuotasTotal + s.capVencido + s.intVencidos + s.moratorios + moratoriosProy + ajustes);
  const avisos = [];
  if (!tasaInfo && s.capVencido > 0) avisos.push('La línea no tiene tasa en el catálogo: los moratorios proyectados quedaron en $0.');
  if (!cuotas.length) avisos.push('No hay cuotas del Rep1 que venzan hasta esa fecha; el cálculo sólo incluye el saldo vencido.');
  if (corteFecha && diffDays_(corteFecha, hoyKey_()) > 0) avisos.push('El Rep9 es del ' + fechaDMA_(corteFecha) + '; los pagos posteriores a ese corte no están reflejados salvo que existan como ajustes.');
  return {
    ok: true, linea: linea, fecha: fechaKey, corteFecha: corteFecha, dias: dias, moneda: String(x[REP9.MONEDA] || 'MXN').trim().toUpperCase() || 'MXN',
    cuotas: cuotas, tasaMoratoria: tasaMoratoria,
    desglose: { cuotas: cuotasTotal, capVencido: s.capVencido, intVencidos: s.intVencidos, moratoriosAcum: s.moratorios, moratoriosProy: moratoriosProy, ajustes: ajustes, total: total },
    saldoTotalRep9: s.saldoTotal, avisos: avisos
  };
}

/** Números de la pantalla Hoy: cartera del Rep9 y movimiento de avisos. */
function armarResumenInicio_() {
  const r9 = rep9Cargado_();
  const cartera = {};
  if (r9) {
    r9.mapa.forEach(x => {
      const m = String(x[REP9.MONEDA] || 'MXN').trim().toUpperCase() || 'MXN';
      const s = saldosDeRep9_(x);
      const t = cartera[m] || (cartera[m] = { saldoTotal: 0, saldoVencido: 0, lineas: 0, conVencido: 0 });
      t.saldoTotal += s.saldoTotal; t.saldoVencido += s.saldoVencido; t.lineas++;
      if (s.capVencido > 0.005 || s.saldoVencido > 0.005) t.conVencido++;
    });
    Object.keys(cartera).forEach(m => { cartera[m].saldoTotal = round2_(cartera[m].saldoTotal); cartera[m].saldoVencido = round2_(cartera[m].saldoVencido); });
  }
  const hoy = hoyKey_(), desde = addDays_(hoy, -6);
  const regs = bitacoraRegistros_(1500).filter(b => b.status === 'ENVIADO' && b.tipo !== 'REENVIO');
  return {
    ok: true, cartera: cartera, hayRep9: !!r9,
    avisosHoy: regs.filter(b => b.timestamp.slice(0, 10) === hoy).length,
    avisosSemana: regs.filter(b => b.timestamp.slice(0, 10) >= desde).length
  };
}


// ─── HOJA DE TRABAJO (réplica de la hoja de cálculo, con su rastro) ─────────

/** Una fila por cuota del Rep1 con las mismas columnas de la hoja de Sheets. */
function armarHojaTrabajo_() {
  const cola = calcularCola_();
  if (!cola.ok) return cola;
  const filas = cola.items.map(i => {
    const t = i.traza, d = i.desglose;
    return {
      key: i.key, linea: i.linea, cliente: i.cliente, moneda: i.moneda, fechaNominal: i.fechaNominal, fechaPago: i.fechaPago,
      capital: d.capital, intereses: d.intereses, otros: d.otros, iva: d.iva, importeRep1: t.rep1.importe,
      capVencido: d.capVencido, intVencidos: d.intVencidos, sumaMoratorios: d.moratoriosAcum, moratoriosPeriodo: d.moratoriosProy,
      dias: i.diasProy, tasaMoratoria: i.tasaMoratoria, ajustes: d.ajustes, total: d.total,
      estado: i.estado, accion: i.accion, sinRep9: i.sinRep9, sinTasa: i.sinTasa, aplicaVencido: t.aplicaVencido,
      cuadra: Math.abs(d.cuota - t.rep1.importe) <= 0.01
    };
  });
  const totales = {};
  filas.forEach(f => {
    const t = totales[f.moneda] || (totales[f.moneda] = { filas: 0, capital: 0, intereses: 0, otros: 0, iva: 0, importeRep1: 0, capVencido: 0, intVencidos: 0, sumaMoratorios: 0, moratoriosPeriodo: 0, ajustes: 0, total: 0 });
    t.filas++;
    ['capital', 'intereses', 'otros', 'iva', 'importeRep1', 'capVencido', 'intVencidos', 'sumaMoratorios', 'moratoriosPeriodo', 'ajustes', 'total'].forEach(k => { t[k] = round2_(t[k] + f[k]); });
  });
  return {
    ok: true, hoy: cola.hoy, corte1: cola.frescura.rep1.corte, corte9: cola.frescura.rep9.corte,
    factor: cfgNum_('FACTOR_TASA_MORATORIA', 2), base: cfgStr_('BASE_MORATORIOS', 'NOMINAL').toUpperCase() === 'EFECTIVA' ? 'EFECTIVA' : 'NOMINAL',
    baseDias: BASE_DIAS_ANIO, filas: filas, totales: totales
  };
}

/** Rep9 con las dos columnas calculadas de la hoja "Saldos_Vencidos" (AE y AF). */
function armarSaldosVencidos_() {
  const r9 = rep9Cargado_();
  if (!r9) return { ok: false, codigo: 'SIN_REP9', error: 'Todavía no hay un Rep9 cargado.' };
  const filas = [], totales = {};
  r9.rows.forEach(x => {
    const linea = normLinea_(x[REP9.LINEA]); if (!linea) return;
    const moneda = String(x[REP9.MONEDA] || 'MXN').trim().toUpperCase() || 'MXN';
    const f = {
      linea: linea, cliente: String(x[REP9.NOMBRE] || ''), moneda: moneda,
      capVigente: round2_(num_(x[REP9.CAP_VIGENTE])), capVencido: round2_(num_(x[REP9.CAP_VENCIDO])),
      intVencido: round2_(num_(x[REP9.INT_VENCIDO])), ivaIntVencido: round2_(num_(x[REP9.IVA_INT_VENCIDO])),
      moratorios: round2_(num_(x[REP9.MORATORIOS])), ivaMoratorios: round2_(num_(x[REP9.IVA_MORATORIOS])),
      morCont: round2_(num_(x[REP9.MOR_CONT])), ivaMorCont: round2_(num_(x[REP9.IVA_MOR_CONT])),
      saldoVencido: round2_(num_(x[REP9.SALDO_VENCIDO])), saldoTotal: round2_(num_(x[REP9.SALDO_TOTAL]))
    };
    f.sumaMoratorios = round2_(f.moratorios + f.ivaMoratorios + f.morCont + f.ivaMorCont);
    f.interesesVencidos = round2_(f.intVencido + f.ivaIntVencido);
    filas.push(f);
    const t = totales[moneda] || (totales[moneda] = { filas: 0, capVigente: 0, capVencido: 0, intVencido: 0, ivaIntVencido: 0, moratorios: 0, ivaMoratorios: 0, morCont: 0, ivaMorCont: 0, sumaMoratorios: 0, interesesVencidos: 0, saldoVencido: 0, saldoTotal: 0 });
    t.filas++;
    Object.keys(t).forEach(k => { if (k !== 'filas') t[k] = round2_(t[k] + f[k]); });
  });
  return { ok: true, corte: r9.corte, filas: filas, totales: totales };
}

/** El reporte tal como quedó guardado (para verlo dentro de la plataforma). */
function leerReporteCargado_(tipo) {
  const esRep1 = tipo === 'rep1';
  if (!esRep1 && tipo !== 'rep9') return { ok: false, error: 'Reporte desconocido.' };
  const sh = spreadsheet_().getSheetByName(esRep1 ? SHEETS.CACHE_REP1 : SHEETS.CACHE_REP9);
  if (!sh || sh.getLastRow() < 5) return { ok: false, codigo: esRep1 ? 'SIN_REP1' : 'SIN_REP9', error: 'Todavía no hay un ' + (esRep1 ? 'Rep1' : 'Rep9') + ' cargado.' };
  const n = esRep1 ? 9 : 30;
  const filas = leerCache_(sh, n).map(r => r.map(c => {
    if (Object.prototype.toString.call(c) === '[object Date]') return fechaKeyDeHoja_(c) || '';
    return (c === null || c === undefined) ? '' : c;
  }));
  return { ok: true, tipo: tipo, encabezados: esRep1 ? REP1_HEADERS : REP9_HEADERS, filas: filas, corte: leerCorteMeta_(sh) };
}

/** Historial de cargas (hoja Cortes), de la más reciente a la más antigua. */
function leerCortes_(limite) {
  const sh = spreadsheet_().getSheetByName(SHEETS.CORTES);
  if (!sh || sh.getLastRow() < 2) return [];
  const n = Math.min(limite || 60, sh.getLastRow() - 1);
  const ini = sh.getLastRow() - n + 1;
  return sh.getRange(ini, 1, n, 10).getValues().map(r => {
    let ts = r[2];
    if (Object.prototype.toString.call(ts) === '[object Date]') ts = Utilities.formatDate(ts, TZ, 'yyyy-MM-dd HH:mm:ss');
    return { corte: String(r[0] || ''), reporte: String(r[1] || ''), fechaHora: String(ts || '').replace(/^'/, ''), usuario: String(r[3] || ''),
             archivo: String(r[4] || ''), filas: Number(r[5]) || 0, vencMin: fechaKeyDeHoja_(r[6]) || '', vencMax: fechaKeyDeHoja_(r[7]) || '',
             rango: String(r[8] || ''), control: num_(r[9]) };
  }).filter(c => c.corte).reverse();
}
