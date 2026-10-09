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
