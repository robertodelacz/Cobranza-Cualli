/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Chat.gs
 *  Recordatorios y resúmenes en un espacio de Google Chat (webhook entrante).
 * ═══════════════════════════════════════════════════════════════════════════
 *  El URL del webhook NO va en el código ni en la hoja: se guarda en las
 *  Propiedades del script (Configuración > Chat, o guardarWebhookChat).
 */

function webhookChat_() {
  return PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL') || '';
}

function urlApp_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function mencionChat_() {
  const m = cfgStr_('CHAT_MENCION', '');
  if (!m) return '';
  if (m.toLowerCase() === 'all') return '<users/all> ';
  return '<' + (m.indexOf('users/') === 0 ? m : 'users/' + m) + '> ';
}

/** Publica un mensaje de texto. Devuelve { ok, codigo, error }. */
function chatEnviar_(texto, conMencion) {
  const url = webhookChat_();
  if (!url) return { ok: false, error: 'No hay webhook de Chat configurado.' };
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
      payload: JSON.stringify({ text: (conMencion ? mencionChat_() : '') + texto })
    });
    const code = resp.getResponseCode();
    return code >= 200 && code < 300 ? { ok: true, codigo: code } : { ok: false, codigo: code, error: 'Chat respondió ' + code };
  } catch (e) {
    return { ok: false, error: String(e.message || e) };
  }
}

function fechaChat_(k) {
  const d = ['dom', 'lun', 'mar', 'mié', 'jue', 'vie', 'sáb'][dow_(k)];
  return d + ' ' + Number(k.slice(8, 10)) + '-' + MESES_CORTO[Number(k.slice(5, 7)) - 1];
}

// ─── Recordatorios (triggers diarios) ──────────────────────────────────────

/** Mensaje de la mañana: qué reportes bajar hoy y con qué rango. */
function recordatorioDescarga() {
  const t = estadoTanda_();
  if (!t.esHabil) return;
  const url = urlApp_();
  const v = t.ventana;
  let msg;
  if (t.inicioTanda) {
    msg = '*Hoy toca descargar los reportes de cobranza*\n' +
      '• *Rep1* con vencimientos del *' + fechaDMA_(v.nominalDesde) + '* al *' + fechaDMA_(v.nominalHasta) + '* (' + fechaChat_(v.nominalDesde) + ' a ' + fechaChat_(v.nominalHasta) + ').\n' +
      '• *Rep9* con corte de hoy (' + fechaChat_(t.hoy) + ').\n' +
      'Cárgalos en la plataforma' + (url ? ': ' + url : '.');
  } else if (cfgNum_('MAX_EDAD_REP9_DIAS_HABILES', 0) === 0) {
    msg = '*Hoy no inicia tanda, pero falta el Rep9 con corte de hoy* (' + fechaChat_(t.hoy) + ').' + (url ? '\n' + url : '');
  } else return;
  chatEnviar_(msg, true);
}

/** Segundo aviso: si ya pasó la hora límite y los reportes no están cargados. */
function escalarDescarga() {
  const t = estadoTanda_();
  if (!t.esHabil) return;
  const ss = spreadsheet_();
  const sh1 = ss.getSheetByName(SHEETS.CACHE_REP1), sh9 = ss.getSheetByName(SHEETS.CACHE_REP9);
  const c1 = sh1 ? leerCorteMeta_(sh1) : { fechaKey: null }, c9 = sh9 ? leerCorteMeta_(sh9) : { fechaKey: null };
  const faltan = [];
  if (t.inicioTanda && c1.fechaKey !== t.hoy) faltan.push('Rep1 (vencimientos del ' + fechaDMA_(t.ventana.nominalDesde) + ' al ' + fechaDMA_(t.ventana.nominalHasta) + ')');
  if (cfgNum_('MAX_EDAD_REP9_DIAS_HABILES', 0) === 0 && c9.fechaKey !== t.hoy) faltan.push('Rep9 con corte de hoy');
  if (!faltan.length) return;
  const url = urlApp_();
  chatEnviar_('*Falta cargar reportes y hoy hay avisos por enviar*\n• ' + faltan.join('\n• ') + (url ? '\n' + url : ''), true);
}

/** Resumen posterior a un envío (lo llama la plataforma y el envío automático). */
function chatResumenEnvio_(res, pendientesRevision, origen) {
  const lineas = ['*Avisos de cobranza (' + (origen || 'manual') + ')*',
    '• Enviados: *' + (res.enviados || 0) + '*' + (res.errores ? '  • Con error: *' + res.errores + '*' : '') + (res.omitidos ? '  • Omitidos: ' + res.omitidos : '')];
  if (res.pendientes && res.pendientes.length) lineas.push('• Quedaron *' + res.pendientes.length + '* por enviar (tiempo o cuota diaria agotados).');
  if (pendientesRevision && pendientesRevision.length) lineas.push('• Esperan revisión o tienen bloqueo: *' + pendientesRevision.length + '*. Revísalas en la plataforma.');
  if (cfgBool_('MODO_PRUEBA', false)) lineas.push('_Modo prueba: los correos fueron a un buzón interno._');
  return chatEnviar_(lineas.join('\n'), false);
}

// ─── Triggers ──────────────────────────────────────────────────────────────

/**
 * Crea (o recrea) los triggers diarios. Ejecutar una vez desde el editor o
 * desde Configuración. Apps Script lanza el trigger en algún momento de la
 * hora indicada.
 */
function instalarTriggers() {
  const manejadores = ['recordatorioDescarga', 'escalarDescarga', 'envioAutomaticoDiario', 'cronEnvioDiario'];
  ScriptApp.getProjectTriggers().forEach(t => { if (manejadores.indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  const creados = [];
  const crear = (fn, hora) => { ScriptApp.newTrigger(fn).timeBased().everyDays(1).atHour(hora).nearMinute(5).inTimezone(TZ).create(); creados.push(fn + ' ~' + hora + ':05'); };
  crear('recordatorioDescarga', Math.round(cfgNum_('HORA_RECORDATORIO', 8)));
  crear('escalarDescarga', Math.round(cfgNum_('HORA_ESCALACION', 9)));
  if (cfgBool_('ENVIO_AUTOMATICO', false)) crear('envioAutomaticoDiario', Math.round(cfgNum_('HORA_ENVIO_AUTOMATICO', 10)));
  return { ok: true, triggers: creados };
}

function triggersActivos_() {
  return ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
}
