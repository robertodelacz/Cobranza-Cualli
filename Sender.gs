/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Sender.gs
 *  Envío en lote, reenvío, vista previa y envío automático.
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  - El motor se calcula UNA vez por llamada (la v1 lo recalculaba por correo).
 *  - Un candado impide que dos personas envíen al mismo tiempo.
 *  - Se respeta el límite de 6 min de Apps Script: si se agota el presupuesto
 *    de tiempo, lo que falte se devuelve como "pendientes" y se retoma después
 *    (el estado sale de la bitácora, así que nada se envía dos veces).
 *  - En MODO_PRUEBA todo llega a un buzón de prueba, se marca ENVIADO_PRUEBA
 *    y NO consume el aviso real.
 */

const PRESUPUESTO_MS = 240000;

function destinoReal_(item) {
  if (cfgBool_('MODO_PRUEBA', false)) {
    const d = cfgStr_('MODO_PRUEBA_DESTINO', '');
    const buzon = d || usuarioActual_().email;
    return { prueba: true, to: [buzon], cc: '', bcc: '' };
  }
  return { prueba: false, to: item.destinatarios, cc: cfgStr_('CC_FIJO', ''), bcc: cfgStr_('BCC_FIJO', '') };
}

function asuntoCodificado_(asunto) {
  if (!cfgBool_('ASUNTO_RFC2047', true)) return asunto;
  return '=?UTF-8?B?' + Utilities.base64Encode(Utilities.newBlob(asunto).getBytes()) + '?=';
}

function replyTo_() {
  return cfgStr_('REPLY_TO', '') || cfgStr_('NOTIFICAR_A', '') || usuarioActual_().email;
}

// ─── VISTA PREVIA ──────────────────────────────────────────────────────────

function previsualizar_(key) {
  const calc = calcularCola_();
  if (!calc.ok) return calc;
  const item = calc.items.find(i => i.key === key);
  if (!item) return { ok: false, error: 'La cuota ya no aparece en los reportes cargados.' };
  const reg = leerBitacoraMapa_().get(key);
  const mail = construirCorreo_(item, { montoAnterior: reg && reg.ultimo ? reg.ultimo.total : null, esReenvio: false });
  const d = destinoReal_(item);
  return { ok: true, asunto: mail.asunto, html: mail.html, para: item.destinatarios, cc: cfgStr_('CC_FIJO', ''),
           desde: usuarioActual_().email, responderA: replyTo_(), pruebaDestino: d.prueba ? d.to[0] : '',
           item: item };
}

// ─── ENVÍO ─────────────────────────────────────────────────────────────────

/**
 * Envía las cuotas indicadas. opts: { origen: 'MANUAL'|'AUTO'|'REENVIO', confirmarRevision, motivo }.
 * Devuelve { ok, enviados, errores, omitidos, pendientes[], resultados[] }.
 */
function ejecutarEnvios_(keys, opts) {
  opts = opts || {};
  const origen = opts.origen || 'MANUAL';
  const esReenvio = origen === 'REENVIO';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, error: 'Hay otro envío en curso. Espera a que termine e intenta de nuevo.' };

  const inicio = Date.now();
  const bitacora = [];
  const resultados = [];
  const pendientes = [];
  let enviados = 0, errores = 0, omitidos = 0;

  try {
    const calc = calcularCola_();
    if (!calc.ok) return calc;
    const mapa = new Map(calc.items.map(i => [i.key, i]));
    const reg = leerBitacoraMapa_();
    const usuario = usuarioActual_().email;
    const ahora = new Date();
    let cuotaDisponible = Infinity;
    try { cuotaDisponible = MailApp.getRemainingDailyQuota(); } catch (e) {}

    for (let n = 0; n < keys.length; n++) {
      const key = keys[n];
      if (Date.now() - inicio > PRESUPUESTO_MS) { pendientes.push.apply(pendientes, keys.slice(n)); break; }
      const item = mapa.get(key);
      const omitir = (motivo) => {
        omitidos++;
        resultados.push({ key: key, ok: false, estado: 'OMITIDO', error: motivo });
        if (item) bitacora.push(filaBitacora_(item, ahora, usuario, calc, 'OMITIDO', motivo, origen, (item.accion || 'AVISO'), [], ''));
      };
      if (!item) { omitidos++; resultados.push({ key: key, ok: false, estado: 'OMITIDO', error: 'La cuota ya no aparece en los reportes.' }); continue; }

      if (esReenvio) {
        if (item.diasHabiles < 0 || item.estado === 'VENCIDA') { omitir('La cuota ya venció; no se reenvía un aviso preventivo.'); continue; }
        if (item.bloqueos.length) { omitir(item.bloqueos.map(b => b.texto).join(' ')); continue; }
      } else {
        if (!item.accion) { omitir('Hoy no le toca aviso (estado: ' + item.estado + ').'); continue; }
        if (item.bloqueos.length) { omitir(item.bloqueos.map(b => b.texto).join(' ')); continue; }
        if (item.estado === 'REVISAR' && !opts.confirmarRevision) { omitir('Tiene alertas por revisar; confírmalas para enviarla.'); continue; }
      }

      const destino = destinoReal_(item);
      const necesarios = destino.to.length + (destino.cc ? destino.cc.split(',').length : 0) + (destino.bcc ? destino.bcc.split(',').length : 0);
      if (cuotaDisponible < necesarios) { pendientes.push.apply(pendientes, keys.slice(n)); resultados.push({ key: key, ok: false, estado: 'PENDIENTE', error: 'Se agotó la cuota diaria de correos de Google.' }); break; }

      const previo = reg.get(key);
      const mail = construirCorreo_(item, { montoAnterior: previo && previo.ultimo ? previo.ultimo.total : null, esReenvio: esReenvio });
      const tipo = esReenvio ? 'REENVIO' : item.accion;
      try {
        const asunto = (destino.prueba ? '[PRUEBA] ' : '') + mail.asunto;
        const advanced = { htmlBody: mail.html, name: cfgStr_('REMITENTE_NOMBRE', 'Cobranza Cualli'), replyTo: replyTo_() };
        if (destino.cc) advanced.cc = destino.cc;
        if (destino.bcc) advanced.bcc = destino.bcc;
        MailApp.sendEmail(destino.to.join(','), asuntoCodificado_(asunto), mail.plain, advanced);
        cuotaDisponible -= necesarios;
        enviados++;
        const status = destino.prueba ? 'ENVIADO_PRUEBA' : 'ENVIADO';
        const msg = (esReenvio ? 'Reenvío: ' + (opts.motivo || '') : 'OK') + (destino.prueba ? ' (modo prueba → ' + destino.to[0] + ')' : '');
        bitacora.push(filaBitacora_(item, ahora, usuario, calc, status, msg, origen, tipo, destino.to, destino.prueba ? destino.to[0] : item.destinatarios.join(', ')));
        resultados.push({ key: key, ok: true, estado: status, cliente: item.cliente });
        Utilities.sleep(150);
      } catch (err) {
        errores++;
        bitacora.push(filaBitacora_(item, ahora, usuario, calc, 'ERROR', String(err.message || err), origen, tipo, destino.to, item.destinatarios.join(', ')));
        resultados.push({ key: key, ok: false, estado: 'ERROR', error: String(err.message || err), cliente: item.cliente });
      }
      if (bitacora.length >= 10) { agregarBitacora_(bitacora.splice(0, bitacora.length)); }
    }
  } finally {
    try { if (bitacora.length) agregarBitacora_(bitacora); } finally { lock.releaseLock(); }
  }
  return { ok: true, enviados: enviados, errores: errores, omitidos: omitidos, pendientes: pendientes, resultados: resultados,
           duracionSeg: Math.round((Date.now() - inicio) / 1000) };
}

function filaBitacora_(item, ahora, usuario, calc, status, mensaje, origen, tipo, destinatarios, correosTexto) {
  const d = item.desglose;
  const id = 'AV-' + item.linea + '-' + item.fechaNominal.replace(/-/g, '') + '-' + Utilities.formatDate(ahora, TZ, 'yyyyMMddHHmmss');
  return [ahora, fechaDeKey_(item.fechaNominal), tipo, item.linea, item.cliente,
          correosTexto || (destinatarios || []).join(', '), item.total, status, mensaje, usuario,
          calc.frescura.rep1.corte.corteId, calc.frescura.rep9.corte.corteId, item.moneda, item.diasHabiles,
          d.cuota, round2_(d.capVencido + d.intVencidos), round2_(d.moratoriosAcum + d.moratoriosProy), origen, id, fechaDeKey_(item.fechaPago)];
}

// ─── ENVÍO AUTOMÁTICO (trigger) ────────────────────────────────────────────

function envioAutomaticoDiario() {
  if (!cfgBool_('ENVIO_AUTOMATICO', false)) return;
  const hoy = hoyKey_();
  if (!esHabil_(hoy)) return;
  const calc = calcularCola_();
  if (!calc.ok) { chatEnviar_('⚠️ *Envío automático detenido:* ' + calc.error); return; }
  const listas = calc.items.filter(i => i.estado === 'LISTA').map(i => i.key);
  const bloqueadas = calc.items.filter(i => i.estado === 'BLOQUEADA' || i.estado === 'REVISAR');
  if (!listas.length && !bloqueadas.length) return;
  const res = listas.length ? ejecutarEnvios_(listas, { origen: 'AUTO' }) : { ok: true, enviados: 0, errores: 0, omitidos: 0, pendientes: [] };
  chatResumenEnvio_(res, bloqueadas, 'automático');
}
