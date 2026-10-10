/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Motor.gs
 *  Calcula, una sola vez por llamada, la cola completa de avisos.
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  Cada cuota (línea + fecha nominal del Rep1) recibe como máximo dos avisos:
 *    PREVENTIVO  la primera vez que se ve a 5 días hábiles o menos de su pago.
 *    VISPERA     cuando falta 1 día hábil, si ya hubo preventivo y pasaron al
 *                menos SEPARACION_MINIMA_DIAS días hábiles desde entonces.
 *  Lo que ya se envió se lee de la bitácora, así el proceso es idempotente.
 *
 *  Estados de una cuota:
 *    LISTA          toca avisar y no falta nada.
 *    REVISAR        toca avisar, pero hay una alerta que conviene mirar antes.
 *    BLOQUEADA      toca avisar, pero falta un dato o el corte está vencido.
 *    PROGRAMADA     aún no entra a la ventana del aviso 1.
 *    ESPERA_VISPERA ya tuvo aviso 1; falta el aviso 2.
 *    COMPLETA       ya recibió lo que le correspondía.
 *    VENCE_HOY / VENCIDA   salen del flujo preventivo.
 */

/** Antigüedad de cada corte contra el límite configurado (en días hábiles). */
function calcularFrescura_(c1, c9, hoy, L) {
  const max1 = Math.round(cfgNum_('MAX_EDAD_REP1_DIAS_HABILES', L - 1));
  const max9 = Math.round(cfgNum_('MAX_EDAD_REP9_DIAS_HABILES', 0));
  const edad1 = c1 && c1.fechaKey ? Math.max(0, habilesEntre_(c1.fechaKey, hoy)) : 999;
  const edad9 = c9 && c9.fechaKey ? Math.max(0, habilesEntre_(c9.fechaKey, hoy)) : 999;
  return {
    rep1: { corte: c1, edad: edad1, max: max1, ok: edad1 <= max1 },
    rep9: { corte: c9, edad: edad9, max: max9, ok: edad9 <= max9 }
  };
}

function calcularCola_() {
  const hoy = hoyKey_();
  const ss = spreadsheet_();
  const sh1 = ss.getSheetByName(SHEETS.CACHE_REP1);
  const sh9 = ss.getSheetByName(SHEETS.CACHE_REP9);
  if (!sh1 || !sh9) return { ok: false, codigo: 'SIN_HOJAS', error: 'Faltan las hojas de caché. Ejecuta inicializarV2().' };
  if (sh1.getLastRow() < 5) return { ok: false, codigo: 'SIN_REP1', error: 'Todavía no hay un Rep1 cargado.' };
  if (sh9.getLastRow() < 5) return { ok: false, codigo: 'SIN_REP9', error: 'Todavía no hay un Rep9 cargado.' };

  const dias = diasAviso_();
  const L = longitudTanda_();
  const c1 = leerCorteMeta_(sh1), c9 = leerCorteMeta_(sh9);
  const frescura = calcularFrescura_(c1, c9, hoy, L);
  const max1 = frescura.rep1.max, max9 = frescura.rep9.max;

  const factor = cfgNum_('FACTOR_TASA_MORATORIA', 2);
  const baseEfectiva = cfgStr_('BASE_MORATORIOS', 'NOMINAL').toUpperCase() === 'EFECTIVA';

  const rep1 = leerCache_(sh1, 9);
  const rep9 = leerCache_(sh9, 30);
  const tasas = leerTasas_();
  const contactos = leerContactos_();
  const bitacora = leerBitacoraMapa_();

  const r9 = new Map();
  rep9.forEach(r => { const l = normLinea_(r[REP9.LINEA]); if (l && !r9.has(l)) r9.set(l, r); });

  // Ajustes vigentes: activos y capturados sobre el Rep9 que está cargado hoy.
  const ajustesPorLinea = new Map();
  leerAjustes_().forEach(a => {
    if (!a.linea || (a.estado !== 'ACTIVO' && a.estado !== 'PENDIENTE')) return;
    if (a.corte9 !== c9.corteId) return;
    if (!ajustesPorLinea.has(a.linea)) ajustesPorLinea.set(a.linea, []);
    ajustesPorLinea.get(a.linea).push(a);
  });

  // Cuotas del Rep1 con su fecha de pago efectiva.
  const cuotas = [];
  const advertencias = [];
  rep1.forEach(r => {
    const linea = normLinea_(r[REP1.LINEA]);
    if (!linea) return;
    const nominal = fechaKeyDeHoja_(r[REP1.FECHA_VENC]);
    if (!nominal) { advertencias.push('Línea ' + linea + ': fecha de vencimiento inválida.'); return; }
    cuotas.push({ r: r, linea: linea, nominal: nominal, pago: siguienteHabil_(nominal) });
  });
  cuotas.sort((a, b) => a.pago < b.pago ? -1 : a.pago > b.pago ? 1 : (a.nominal < b.nominal ? -1 : a.nominal > b.nominal ? 1 : (a.linea < b.linea ? -1 : 1)));

  const vistas = new Set();
  const lineasConVarias = new Set();
  cuotas.forEach(c => { if (vistas.has(c.linea)) lineasConVarias.add(c.linea); vistas.add(c.linea); });
  const yaConVencido = new Set();

  const items = cuotas.map(c => {
    const r = c.r, linea = c.linea;
    const nombreRep1 = String(r[REP1.NOMBRE] || '');
    const moneda1 = String(r[REP1.MONEDA] || 'MXN').trim().toUpperCase();
    const x9 = r9.get(linea) || null;
    const moneda9 = x9 ? String(x9[REP9.MONEDA] || '').trim().toUpperCase() : '';
    const tasaInfo = tasas.get(linea) || null;
    const contacto = contactos.get(linea) || null;
    const key = linea + '|' + c.nominal;
    const reg = bitacora.get(key) || { previo: null, vispera: null, ultimo: null, reenvios: 0 };

    // ─── Montos ───
    const capital = num_(r[REP1.CAPITAL]), intereses = num_(r[REP1.INTERES]), otros = num_(r[REP1.OTROS]);
    const iva = num_(r[REP1.IVA]), importe = num_(r[REP1.IMPORTE]);
    const cuota = round2_(capital + intereses + otros + iva);

    // El saldo vencido de la línea se cobra una sola vez: en su primera cuota.
    const primera = !yaConVencido.has(linea);
    yaConVencido.add(linea);
    let capVencido = 0, intVencidos = 0, moratoriosAcum = 0;
    if (x9 && primera) {
      capVencido = num_(x9[REP9.CAP_VENCIDO]);
      intVencidos = num_(x9[REP9.INT_VENCIDO]) + num_(x9[REP9.IVA_INT_VENCIDO]);
      moratoriosAcum = num_(x9[REP9.MORATORIOS]) + num_(x9[REP9.IVA_MORATORIOS]) + num_(x9[REP9.MOR_CONT]) + num_(x9[REP9.IVA_MOR_CONT]);
    }
    const tasaContrato = tasaInfo ? tasaInfo.tasa : 0;
    const tasaMoratoria = tasaContrato * factor;
    // Moratorios proyectados desde el corte del Rep9 (no desde hoy) hasta la fecha base.
    const base = baseEfectiva ? c.pago : c.nominal;
    const diasProy = c9.fechaKey ? diffDays_(c9.fechaKey, base) : 0;
    const moratoriosProy = (capVencido > 0 && diasProy > 0 && tasaMoratoria > 0)
      ? capVencido * tasaMoratoria / BASE_DIAS_ANIO * diasProy : 0;

    let ajustes = 0;
    const ajustesAplicados = [];
    let ajustePendiente = false;
    (ajustesPorLinea.get(linea) || []).forEach(a => {
      const aplica = a.vencimiento ? a.vencimiento === c.nominal : primera;
      if (!aplica) return;
      if (a.estado === 'PENDIENTE') { ajustePendiente = true; return; }
      ajustes += a.monto;
      ajustesAplicados.push({ id: a.id, tipo: a.tipo, monto: a.monto, motivo: a.motivo, usuario: a.usuario });
    });

    const total = round2_(cuota + capVencido + intVencidos + moratoriosAcum + moratoriosProy + ajustes);
    const desglose = {
      capital: round2_(capital), intereses: round2_(intereses), otros: round2_(otros), iva: round2_(iva), cuota: cuota,
      capVencido: round2_(capVencido), intVencidos: round2_(intVencidos), moratoriosAcum: round2_(moratoriosAcum),
      moratoriosProy: round2_(moratoriosProy), ajustes: round2_(ajustes), total: total
    };

    // Rastro: de dónde sale cada número (para la hoja de trabajo y el detalle del aviso).
    const traza = {
      rep1: { fecha: c.nominal, capital: round2_(capital), intereses: round2_(intereses), otros: round2_(otros), iva: round2_(iva), importe: round2_(importe) },
      rep9: x9 ? {
        moneda: moneda9, capVigente: round2_(num_(x9[REP9.CAP_VIGENTE])), capVencido: round2_(num_(x9[REP9.CAP_VENCIDO])),
        intVencido: round2_(num_(x9[REP9.INT_VENCIDO])), ivaIntVencido: round2_(num_(x9[REP9.IVA_INT_VENCIDO])),
        moratorios: round2_(num_(x9[REP9.MORATORIOS])), ivaMoratorios: round2_(num_(x9[REP9.IVA_MORATORIOS])),
        morCont: round2_(num_(x9[REP9.MOR_CONT])), ivaMorCont: round2_(num_(x9[REP9.IVA_MOR_CONT])),
        saldoVencido: round2_(num_(x9[REP9.SALDO_VENCIDO])), saldoTotal: round2_(num_(x9[REP9.SALDO_TOTAL]))
      } : null,
      primera: primera, aplicaVencido: !!(x9 && primera),
      tasaContrato: tasaContrato, factor: factor, baseDias: BASE_DIAS_ANIO, base: baseEfectiva ? 'EFECTIVA' : 'NOMINAL',
      fechaBase: base, corte9: c9.fechaKey, diasCrudos: diasProy, tasaEnCatalogo: !!tasaInfo
    };

    // ─── Calendario del aviso ───
    const dh = habilesEntre_(hoy, c.pago);
    let estado, accion = null, programadaPara = null, nota = '';
    if (dh < 0) estado = 'VENCIDA';
    else if (dh === 0) estado = 'VENCE_HOY';
    else if (reg.vispera) estado = 'COMPLETA';
    else if (!reg.previo) {
      if (dh <= dias.a1) accion = 'PREVENTIVO';
      else { estado = 'PROGRAMADA'; programadaPara = restarHabiles_(c.pago, dias.a1); }
    } else if (dh <= dias.a2) {
      if (habilesEntre_(reg.previo.fechaKey, hoy) >= dias.separacion) accion = 'VISPERA';
      else { estado = 'COMPLETA'; nota = 'El aviso 2 se omite: el aviso 1 salió hace muy poco.'; }
    } else { estado = 'ESPERA_VISPERA'; programadaPara = restarHabiles_(c.pago, dias.a2); }

    // ─── Bloqueos y alertas ───
    const bloqueos = [], alertas = [];
    const destinatarios = contacto ? parsearDestinatarios_(contacto.correo) : [];
    if (!contacto || !destinatarios.length) bloqueos.push({ codigo: 'SIN_CORREO', texto: 'No hay un correo válido en el catálogo.' });
    if (!contacto || !contacto.stp) bloqueos.push({ codigo: 'SIN_STP', texto: 'No hay cuenta STP en el catálogo.' });
    if (!frescura.rep1.ok) bloqueos.push({ codigo: 'REP1_VENCIDO', texto: 'El Rep1 es de ' + (c1.fechaKey || 'fecha desconocida') + ' (máximo ' + max1 + ' día(s) hábil(es) de antigüedad).' });
    if (!frescura.rep9.ok) bloqueos.push({ codigo: 'REP9_VENCIDO', texto: 'El Rep9 es de ' + (c9.fechaKey || 'fecha desconocida') + ' (máximo ' + max9 + ' día(s) hábil(es) de antigüedad).' });
    if (x9 && moneda9 && moneda9 !== moneda1) bloqueos.push({ codigo: 'MONEDA_DISTINTA', texto: 'El Rep1 trae ' + moneda1 + ' y el Rep9 ' + moneda9 + '.' });
    if (ajustePendiente) bloqueos.push({ codigo: 'AJUSTE_PENDIENTE', texto: 'Hay un ajuste de saldo esperando aprobación.' });

    if (!x9) alertas.push({ codigo: 'SIN_REP9', nivel: 'warn', texto: 'La línea no aparece en el Rep9: el saldo vencido podría estar incompleto.' });
    if (!tasaInfo) {
      if (capVencido > 0) alertas.push({ codigo: 'SIN_TASA', nivel: 'warn', texto: 'Sin tasa en el catálogo y con capital vencido: los moratorios proyectados quedaron en $0.' });
      else alertas.push({ codigo: 'SIN_TASA', nivel: 'info', texto: 'Sin tasa en el catálogo (no afecta este monto: no hay capital vencido).' });
    }
    if (tasaInfo && tasaInfo.duplicada) alertas.push({ codigo: 'TASA_DUPLICADA', nivel: 'warn', texto: 'La línea está repetida en el catálogo de tasas; se usó la primera.' });
    if (contacto && contacto.duplicado) alertas.push({ codigo: 'CONTACTO_DUPLICADO', nivel: 'warn', texto: 'La línea está repetida en el catálogo de correos; se usó la primera.' });
    if (Math.abs(cuota - importe) > 0.01) alertas.push({ codigo: 'IMPORTE_NO_CUADRA', nivel: 'warn', texto: 'Capital + interés + otros + IVA ($' + cuota.toFixed(2) + ') no coincide con el Importe del Rep1 ($' + importe.toFixed(2) + ').' });
    if (!primera) alertas.push({ codigo: 'VENCIDO_EN_PRIMERA', nivel: 'info', texto: 'La línea tiene otra cuota antes: el saldo vencido va en esa.' });
    if (accion === 'VISPERA' && reg.previo && Math.abs(total - reg.previo.total) > 0.01) {
      alertas.push({ codigo: 'MONTO_CAMBIO', nivel: 'info', texto: 'El monto cambió contra el aviso 1: de $' + reg.previo.total.toFixed(2) + ' a $' + total.toFixed(2) + '.' });
    }
    if (ajustesAplicados.length) alertas.push({ codigo: 'CON_AJUSTES', nivel: 'info', texto: 'Incluye ' + ajustesAplicados.length + ' ajuste(s) de saldo.' });

    if (accion) {
      if (bloqueos.length) estado = 'BLOQUEADA';
      else if (alertas.some(a => a.nivel === 'warn')) estado = 'REVISAR';
      else estado = 'LISTA';
    }

    const tieneVencido = (capVencido + intVencidos + moratoriosAcum + moratoriosProy) > 0.005;
    return {
      key: key, linea: linea, cliente: (contacto && contacto.cliente) || nombreRep1 || (x9 ? String(x9[REP9.NOMBRE] || '') : ''),
      nombreRep1: nombreRep1, moneda: moneda1, fechaNominal: c.nominal, fechaPago: c.pago, diasHabiles: dh,
      estado: estado, accion: accion, programadaPara: programadaPara, nota: nota,
      previo: reg.previo, vispera: reg.vispera, reenvios: reg.reenvios, ultimo: reg.ultimo,
      bloqueos: bloqueos, alertas: alertas, destinatarios: destinatarios, cuentaSTP: contacto ? contacto.stp : '',
      tasaContrato: tasaContrato, tasaMoratoria: tasaMoratoria, diasProy: Math.max(0, diasProy),
      desglose: desglose, traza: traza, ajustes: ajustesAplicados, plantilla: tieneVencido ? 'B' : 'A',
      sinRep9: !x9, sinTasa: !tasaInfo, multiCuota: lineasConVarias.has(linea), total: total
    };
  });

  return { ok: true, hoy: hoy, esHabil: esHabil_(hoy), fechaSimulada: fechaSimulada_(), frescura: frescura,
           dias: dias, longitud: L, items: items, stats: estadisticas_(items), advertencias: advertencias.slice(0, 50) };
}

function estadisticas_(items) {
  const porEstado = {};
  items.forEach(i => porEstado[i.estado] = (porEstado[i.estado] || 0) + 1);
  const porEnviar = items.filter(i => i.accion);
  const suma = (arr, mon) => round2_(arr.filter(i => i.moneda === mon).reduce((s, i) => s + i.total, 0));
  return {
    total: items.length, porEstado: porEstado,
    porEnviar: porEnviar.length,
    aviso1: porEnviar.filter(i => i.accion === 'PREVENTIVO').length,
    aviso2: porEnviar.filter(i => i.accion === 'VISPERA').length,
    sumaMXN: suma(porEnviar, 'MXN'), sumaUSD: suma(porEnviar, 'USD'),
    sinTasa: items.filter(i => i.sinTasa).length, sinRep9: items.filter(i => i.sinRep9).length,
    conVencido: items.filter(i => i.plantilla === 'B').length
  };
}

/** Auditoría de catálogos contra la cartera del Rep9 (lo que falta antes de que estorbe). */
function auditarCatalogos_() {
  const sh9 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP9);
  const rep9 = sh9 ? leerCache_(sh9, 30) : [];
  const tasas = leerTasas_(), contactos = leerContactos_();
  const out = { cartera: rep9.length, sinTasa: [], sinCorreo: [], correoInvalido: [], sinStp: [], tasasDuplicadas: [], contactosDuplicados: [] };
  rep9.forEach(r => {
    const l = normLinea_(r[REP9.LINEA]);
    if (!l) return;
    const nombre = String(r[REP9.NOMBRE] || '');
    const saldo = num_(r[REP9.SALDO_TOTAL]);
    const base = { linea: l, cliente: nombre, saldo: saldo, moneda: String(r[REP9.MONEDA] || '') };
    if (!tasas.has(l)) out.sinTasa.push(base);
    const c = contactos.get(l);
    if (!c) out.sinCorreo.push(base);
    else {
      if (!c.correo) out.sinCorreo.push(base);
      else if (!parsearDestinatarios_(c.correo).length) out.correoInvalido.push(Object.assign({ correo: c.correo }, base));
      if (!c.stp) out.sinStp.push(base);
    }
  });
  tasas.forEach((v, l) => { if (v.duplicada) out.tasasDuplicadas.push({ linea: l }); });
  contactos.forEach((v, l) => { if (v.duplicado) out.contactosDuplicados.push({ linea: l }); });
  return out;
}
