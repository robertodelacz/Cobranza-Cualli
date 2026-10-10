/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Api.gs
 *  Puntos de entrada de la web app (lo que llama el front con google.script.run).
 * ═══════════════════════════════════════════════════════════════════════════
 *  Las funciones públicas devuelven siempre { ok: true, ... } o { ok: false, error }.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cobranza Preventiva · Cualli')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function envolver_(fn) {
  try { return fn(); } catch (err) { return { ok: false, error: String(err && err.message ? err.message : err) }; }
}

function conCandado_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Hay otra operación en curso. Intenta de nuevo en unos segundos.');
  try { return fn(); } finally { lock.releaseLock(); }
}

// ─── ESTADO INICIAL ────────────────────────────────────────────────────────

function getEstadoInicial() {
  return envolver_(() => {
    const u = usuarioActual_();
    const ss = spreadsheet_();
    const sh1 = ss.getSheetByName(SHEETS.CACHE_REP1), sh9 = ss.getSheetByName(SHEETS.CACHE_REP9);
    let triggers = [];
    try { triggers = triggersActivos_(); } catch (e) {}
    const cortes = { rep1: sh1 ? leerCorteMeta_(sh1) : { cargado: false }, rep9: sh9 ? leerCorteMeta_(sh9) : { cargado: false } };
    const fr = calcularFrescura_(cortes.rep1, cortes.rep9, hoyKey_(), longitudTanda_());
    ['rep1', 'rep9'].forEach(k => { if (cortes[k].cargado) { cortes[k].edad = fr[k].edad; cortes[k].max = fr[k].max; cortes[k].ok = fr[k].ok; } });
    return {
      ok: true, version: VERSION,
      usuario: u, permisos: Object.keys(PERMISOS).filter(p => puede_(p)),
      tanda: estadoTanda_(),
      cortes: cortes,
      chatConfigurado: !!webhookChat_(), triggers: triggers,
      modoPrueba: cfgBool_('MODO_PRUEBA', false), fechaSimulada: fechaSimulada_(),
      inicializado: !!(sh1 && sh9 && ss.getSheetByName(SHEETS.CORTES))
    };
  });
}

// ─── CARGA DE REPORTES ─────────────────────────────────────────────────────

function validarReporte(tipo, rows, meta) {
  return envolver_(() => {
    requerir_('cargar');
    const t = estadoTanda_();
    if (tipo === 'rep1') {
      const n = normalizarRep1Raw_(rows);
      if (!n.ok) return n;
      const v = t.ventana, hoy = t.hoy;
      const keys = n.dataRows.map(r => fechaKeyDeHoja_(r[0]));
      const validas = keys.filter(k => k);
      const porFecha = {};
      validas.forEach(k => porFecha[k] = (porFecha[k] || 0) + 1);
      const min = validas.length ? validas.reduce((a, b) => a < b ? a : b) : null;
      const max = validas.length ? validas.reduce((a, b) => a > b ? a : b) : null;
      const avisos = [];
      if (validas.length < keys.length) avisos.push({ nivel: 'warn', texto: (keys.length - validas.length) + ' fila(s) con fecha de vencimiento ilegible; se ignorarán.' });
      if (max && max < addDays_(v.nominalHasta, -2)) avisos.push({ nivel: 'warn', texto: 'El archivo llega hasta el ' + fechaDMA_(max) + ' y esta tanda necesita hasta el ' + fechaDMA_(v.nominalHasta) + '. Confirma que el filtro del reporte fue el recomendado.' });
      if (min && min > addDays_(v.nominalDesde, 2)) avisos.push({ nivel: 'warn', texto: 'El archivo empieza el ' + fechaDMA_(min) + ' y esta tanda necesita desde el ' + fechaDMA_(v.nominalDesde) + '.' });
      const fuera = validas.filter(k => k < v.nominalDesde || k > v.nominalHasta).length;
      if (fuera) avisos.push({ nivel: 'info', texto: fuera + ' cuota(s) quedan fuera del rango recomendado (' + fechaDMA_(v.nominalDesde) + ' a ' + fechaDMA_(v.nominalHasta) + ').' });
      const hoyOAntes = validas.filter(k => siguienteHabil_(k) <= hoy).length;
      if (hoyOAntes) avisos.push({ nivel: 'info', texto: hoyOAntes + ' cuota(s) vencen hoy o antes: no reciben aviso preventivo.' });
      const lineas = n.dataRows.map(r => normLinea_(r[1]));
      const repetidas = lineas.length - new Set(lineas).size;
      if (repetidas) avisos.push({ nivel: 'info', texto: repetidas + ' línea(s) tienen más de una cuota en el archivo; el saldo vencido se cobra solo en la primera.' });
      const monedas = {};
      n.dataRows.forEach(r => monedas[r[8]] = (monedas[r[8]] || 0) + 1);
      return {
        ok: true, tipo: 'rep1', filas: n.dataRows.length, fechaMin: min, fechaMax: max,
        porFecha: Object.keys(porFecha).sort().map(k => ({ fecha: k, habil: esHabil_(k), n: porFecha[k] })),
        requerido: { desde: v.nominalDesde, hasta: v.nominalHasta }, monedas: monedas, avisos: avisos,
        control: round2_(n.dataRows.reduce((s, r) => s + num_(r[7]), 0)), normalizadas: n.allRows.length
      };
    }
    if (tipo === 'rep9') {
      const n = validarRep9Raw_(rows);
      if (!n.ok) return n;
      const avisos = [];
      const lineas = new Set(n.dataRows.map(r => normLinea_(r[0])));
      const sh1 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP1);
      if (sh1 && sh1.getLastRow() >= 5) {
        const faltan = leerCache_(sh1, 9).map(r => normLinea_(r[1])).filter(l => l && !lineas.has(l));
        if (faltan.length) avisos.push({ nivel: 'warn', texto: faltan.length + ' línea(s) del Rep1 cargado no están en este Rep9: ' + Array.from(new Set(faltan)).slice(0, 8).join(', ') + (faltan.length > 8 ? '…' : '') + '.' });
      }
      const monedas = {};
      n.dataRows.forEach(r => { const m = String(r[3] || '—').trim().toUpperCase(); monedas[m] = (monedas[m] || 0) + 1; });
      if (Object.keys(monedas).length === 1 && monedas.MXN) avisos.push({ nivel: 'info', texto: 'Este Rep9 sólo trae líneas en MXN. Si tienes créditos en USD, revisa que el reporte los incluya.' });
      return { ok: true, tipo: 'rep9', filas: n.dataRows.length, monedas: monedas, avisos: avisos,
               control: round2_(n.dataRows.reduce((s, r) => s + num_(r[27]), 0)) };
    }
    return { ok: false, error: 'Tipo de reporte desconocido.' };
  });
}

function guardarReporte(tipo, rows, meta) {
  return envolver_(() => {
    requerir_('cargar');
    return conCandado_(() => {
      let filas;
      if (tipo === 'rep1') { const n = normalizarRep1Raw_(rows); if (!n.ok) return n; filas = n.allRows; }
      else if (tipo === 'rep9') { const n = validarRep9Raw_(rows); if (!n.ok) return n; filas = n.allRows; }
      else return { ok: false, error: 'Tipo de reporte desconocido.' };
      const r = guardarCache_(tipo, filas, meta || {});
      if (r.ok) invalidarConfig_();
      return r;
    });
  });
}

// ─── COLA, VISTA PREVIA Y ENVÍO ────────────────────────────────────────────

function getCola() {
  return envolver_(() => {
    const c = calcularCola_();
    if (c.ok) c.tanda = estadoTanda_();
    return c;
  });
}

function previsualizarAviso(key) { return envolver_(() => previsualizar_(String(key))); }

/** opts: { confirmarRevision: boolean }. El front envía en bloques de ~12 llaves. */
function enviarAvisos(keys, opts) {
  return envolver_(() => {
    requerir_('enviar');
    if (!Array.isArray(keys) || !keys.length) return { ok: false, error: 'No hay cuotas seleccionadas.' };
    if (keys.length > 40) return { ok: false, error: 'Envía como máximo 40 cuotas por llamada.' };
    return ejecutarEnvios_(keys.map(String), { origen: 'MANUAL', confirmarRevision: !!(opts && opts.confirmarRevision) });
  });
}

function reenviarAviso(key, motivo) {
  return envolver_(() => {
    requerir_('enviar');
    const m = String(motivo || '').trim();
    if (m.length < 5) return { ok: false, error: 'Escribe el motivo del reenvío (mínimo 5 caracteres).' };
    return ejecutarEnvios_([String(key)], { origen: 'REENVIO', motivo: m });
  });
}

function resumenEnChat(res, bloqueadas) {
  return envolver_(() => {
    requerir_('enviar');
    return chatResumenEnvio_(res || {}, new Array(Number(bloqueadas) || 0), 'manual');
  });
}

// ─── BITÁCORA ──────────────────────────────────────────────────────────────

function getBitacora(f) {
  return envolver_(() => {
    f = f || {};
    let regs = bitacoraRegistros_(Math.max(50, Math.min(2000, Number(f.limite) || 400)));
    if (f.status) regs = regs.filter(x => x.status === f.status);
    if (f.texto) { const q = String(f.texto).toLowerCase(); regs = regs.filter(x => (x.linea + ' ' + x.cliente + ' ' + x.correos).toLowerCase().indexOf(q) >= 0); }
    if (f.desde) regs = regs.filter(x => x.timestamp.slice(0, 10) >= f.desde);
    return { ok: true, registros: regs.reverse() };
  });
}

// ─── AJUSTES DE SALDO ──────────────────────────────────────────────────────

function getAjustes() {
  return envolver_(() => {
    const sh9 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP9);
    const corte9 = sh9 ? leerCorteMeta_(sh9).corteId : '';
    const lista = leerAjustes_().reverse().slice(0, 300).map(a => Object.assign({ vigente: a.corte9 === corte9 && (a.estado === 'ACTIVO' || a.estado === 'PENDIENTE') }, a));
    const u = usuarioActual_();
    return { ok: true, ajustes: lista, corteRep9: corte9, umbral: cfgNum_('UMBRAL_AJUSTE_APROBACION', 50000), conAprobacion: !!cfgStr_('APROBADORES', ''), yo: u.email };
  });
}

function crearAjuste(a) {
  return envolver_(() => {
    requerir_('ajustar');
    const linea = normLinea_(a && a.linea);
    const motivo = String((a && a.motivo) || '').trim();
    const tipo = String((a && a.tipo) || '');
    if (!linea) return { ok: false, error: 'Indica la línea de crédito.' };
    if (['PAGO_APLICADO', 'DISPOSICION', 'CORRECCION'].indexOf(tipo) < 0) return { ok: false, error: 'Elige el tipo de ajuste.' };
    if (motivo.length < 8) return { ok: false, error: 'Escribe el motivo con un poco más de detalle (mínimo 8 caracteres).' };
    let monto = num_(a.monto);
    if (!monto) return { ok: false, error: 'El monto debe ser distinto de cero.' };
    if (tipo === 'PAGO_APLICADO') monto = -Math.abs(monto);
    if (tipo === 'DISPOSICION') monto = Math.abs(monto);
    const sh9 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP9);
    const corte9 = sh9 ? leerCorteMeta_(sh9).corteId : '';
    if (!corte9) return { ok: false, error: 'Primero carga un Rep9: el ajuste se liga a ese corte.' };
    const existe = sh9 && leerCache_(sh9, 1).some(r => normLinea_(r[0]) === linea);
    if (!existe) return { ok: false, error: 'La línea ' + linea + ' no está en el Rep9 cargado.' };
    const venc = a.vencimiento ? fechaKeyDeHoja_(a.vencimiento) : '';
    const aprobadores = cfgStr_('APROBADORES', '');
    const pendiente = !!aprobadores && Math.abs(monto) >= cfgNum_('UMBRAL_AJUSTE_APROBACION', 50000);
    const u = usuarioActual_();
    return conCandado_(() => {
      const sh = spreadsheet_().getSheetByName(SHEETS.AJUSTES);
      if (!sh) return { ok: false, error: 'Falta la hoja Ajustes. Ejecuta inicializarV2().' };
      const id = 'AJ-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
      const r = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(r, 1, 1, 12).setValues([[id, new Date(), u.email, linea, venc ? fechaDeKey_(venc) : '', tipo, monto, motivo, pendiente ? 'PENDIENTE' : 'ACTIVO', corte9, '', '']]);
      sh.getRange(r, 2).setNumberFormat('yyyy-mm-dd hh:mm'); sh.getRange(r, 5).setNumberFormat('dd/mm/yyyy'); sh.getRange(r, 7).setNumberFormat('$#,##0.00');
      return { ok: true, id: id, estado: pendiente ? 'PENDIENTE' : 'ACTIVO', monto: monto };
    });
  });
}

function cambiarEstadoAjuste_(id, nuevo, nota) {
  const sh = spreadsheet_().getSheetByName(SHEETS.AJUSTES);
  const a = leerAjustes_().find(x => x.id === id);
  if (!a) throw new Error('No se encontró el ajuste ' + id + '.');
  const u = usuarioActual_();
  if (nuevo === 'ACTIVO') {
    if (a.estado !== 'PENDIENTE') throw new Error('El ajuste no está pendiente de aprobación.');
    if (a.usuario.toLowerCase() === u.email) throw new Error('Quien captura el ajuste no puede aprobarlo.');
    const lista = cfgStr_('APROBADORES', '').toLowerCase().split(/[,;]/).map(s => s.trim()).filter(Boolean);
    if (lista.length && lista.indexOf(u.email) < 0 && u.rol !== 'ADMIN') throw new Error('Tu cuenta no está en la lista de aprobadores.');
    sh.getRange(a.fila, 9).setValue('ACTIVO'); sh.getRange(a.fila, 11).setValue(u.email);
  } else {
    sh.getRange(a.fila, 9).setValue('ANULADO');
    sh.getRange(a.fila, 12).setValue('Anulado por ' + u.email + ' (' + ahoraMX_() + '): ' + nota);
  }
}

function aprobarAjuste(id) {
  return envolver_(() => { requerir_('aprobar'); return conCandado_(() => { cambiarEstadoAjuste_(String(id), 'ACTIVO'); return { ok: true }; }); });
}

function anularAjuste(id, motivo) {
  return envolver_(() => {
    requerir_('ajustar');
    const m = String(motivo || '').trim();
    if (m.length < 5) return { ok: false, error: 'Escribe por qué se anula (mínimo 5 caracteres).' };
    return conCandado_(() => { cambiarEstadoAjuste_(String(id), 'ANULADO', m); return { ok: true }; });
  });
}

// ─── CATÁLOGOS ─────────────────────────────────────────────────────────────

function getCatalogos() {
  return envolver_(() => {
    const a = auditarCatalogos_();
    return { ok: true, auditoria: a, puedeTasas: puede_('tasas'), puedeContactos: puede_('catalogo') };
  });
}

function getFichaLinea(linea) {
  return envolver_(() => {
    const l = normLinea_(linea);
    if (!l) return { ok: false, error: 'Indica la línea.' };
    const t = leerTasas_().get(l), c = leerContactos_().get(l);
    const sh9 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP9);
    const r9 = sh9 ? leerCache_(sh9, 30).find(r => normLinea_(r[0]) === l) : null;
    return { ok: true, linea: l, cliente: (c && c.cliente) || (t && t.nombre) || (r9 ? r9[1] : ''), enRep9: !!r9,
             tasa: t ? t.tasa : null, correo: c ? c.correo : '', stp: c ? c.stp : '', numCliente: c ? c.numCliente : (r9 ? r9[2] : '') };
  });
}

function registrarCambio_(hoja, linea, campo, antes, despues) {
  const sh = spreadsheet_().getSheetByName(SHEETS.CAMBIOS);
  if (!sh) return;
  sh.getRange(Math.max(sh.getLastRow(), 1) + 1, 1, 1, 7).setValues([[new Date(), usuarioActual_().email, hoja, linea, campo, antes, despues]]);
}

function guardarTasa(linea, nombre, tasa) {
  return envolver_(() => {
    requerir_('tasas');
    const l = normLinea_(linea);
    let t = num_(tasa);
    if (t > 1.5) t = t / 100;               // 36 → 0.36
    if (!l || !(t > 0 && t <= 1.5)) return { ok: false, error: 'Indica la línea y una tasa válida (por ejemplo 36% o 0.36).' };
    return conCandado_(() => {
      const sh = spreadsheet_().getSheetByName(SHEETS.TASAS);
      const previa = leerTasas_().get(l);
      if (previa) {
        sh.getRange(previa.fila, 3).setValue(t);
        registrarCambio_('Tasas', l, 'Tasa', previa.tasa, t);
      } else {
        const r = Math.max(sh.getLastRow(), 2) + 1;
        sh.getRange(r, 1, 1, 3).setValues([[Number(l), String(nombre || ''), t]]);
        registrarCambio_('Tasas', l, 'Alta', '', t);
      }
      return { ok: true, tasa: t };
    });
  });
}

function guardarContacto(linea, cliente, correos, stp) {
  return envolver_(() => {
    requerir_('catalogo');
    const l = normLinea_(linea);
    const lista = parsearDestinatarios_(correos);
    const raw = String(correos || '').split(/[,;]/).map(s => s.trim()).filter(Boolean);
    const cuenta = String(stp || '').replace(/\s/g, '');
    if (!l) return { ok: false, error: 'Indica la línea.' };
    if (!lista.length || lista.length !== raw.length) return { ok: false, error: 'Revisa los correos: todos deben ser válidos y estar separados por coma.' };
    if (!/^\d{18}$/.test(cuenta)) return { ok: false, error: 'La cuenta STP debe tener 18 dígitos.' };
    return conCandado_(() => {
      const sh = spreadsheet_().getSheetByName(SHEETS.CORREOS);
      const previo = leerContactos_().get(l);
      if (previo) {
        sh.getRange(previo.fila, 5).setNumberFormat('@');
        sh.getRange(previo.fila, 3, 1, 3).setValues([[cliente || previo.cliente, lista.join(', '), cuenta]]);
        registrarCambio_('Correos', l, 'Correos / STP', previo.correo + ' | ' + previo.stp, lista.join(', ') + ' | ' + cuenta);
      } else {
        const sh9 = spreadsheet_().getSheetByName(SHEETS.CACHE_REP9);
        const r9 = sh9 ? leerCache_(sh9, 30).find(r => normLinea_(r[0]) === l) : null;
        const r = Math.max(sh.getLastRow(), 2) + 1;
        sh.getRange(r, 5).setNumberFormat('@');
        sh.getRange(r, 1, 1, 5).setValues([[r9 ? r9[2] : '', Number(l), cliente || (r9 ? r9[1] : ''), lista.join(', '), cuenta]]);
        registrarCambio_('Correos', l, 'Alta', '', lista.join(', ') + ' | ' + cuenta);
      }
      return { ok: true };
    });
  });
}

// ─── CONFIGURACIÓN ─────────────────────────────────────────────────────────

const VALIDADORES_CONFIG = {
  TANDA_DIAS_HABILES: v => /^\d+$/.test(v) && Number(v) >= 1 && Number(v) <= 10,
  BASE_MORATORIOS: v => ['NOMINAL', 'EFECTIVA'].indexOf(String(v).toUpperCase()) >= 0,
  HORA_RECORDATORIO: v => /^\d+$/.test(v) && Number(v) <= 23,
  HORA_ESCALACION: v => /^\d+$/.test(v) && Number(v) <= 23,
  HORA_ENVIO_AUTOMATICO: v => /^\d+$/.test(v) && Number(v) <= 23,
  MAX_EDAD_REP9_DIAS_HABILES: v => /^\d+$/.test(v) && Number(v) <= 10,
  MAX_EDAD_REP1_DIAS_HABILES: v => v === '' || (/^\d+$/.test(v) && Number(v) <= 10),
  DIAS_TIPO_T_MENOS_5: v => /^\d+$/.test(v) && Number(v) >= 1 && Number(v) <= 15,
  DIAS_TIPO_T_MENOS_1: v => /^\d+$/.test(v) && Number(v) >= 1 && Number(v) <= 15,
  TANDA_ANCHOR_FECHA: v => v === '' || /^\d{4}-\d{2}-\d{2}$/.test(v)
};

function getConfigUI() {
  return envolver_(() => {
    const sh = spreadsheet_().getSheetByName(SHEETS.CONFIG);
    const filas = sh && sh.getLastRow() >= 3 ? sh.getRange(3, 1, sh.getLastRow() - 2, 3).getValues() : [];
    const params = filas.filter(r => String(r[0]).trim()).map(r => ({
      clave: String(r[0]).trim(), valor: r[1] === true ? 'TRUE' : r[1] === false ? 'FALSE' : String(r[1] === null || r[1] === undefined ? '' : r[1]),
      descripcion: String(r[2] || ''), sensible: /^CUENTA_/.test(String(r[0]))
    }));
    return { ok: true, params: params, inhabiles: inhabilesDelAnio_(Number(hoyKey_().slice(0, 4))).concat(inhabilesDelAnio_(Number(hoyKey_().slice(0, 4)) + 1)),
             chatConfigurado: !!webhookChat_(), triggers: triggersActivos_(), puedeConfig: puede_('config') };
  });
}

function setConfig(clave, valor) {
  return envolver_(() => {
    requerir_('config');
    const k = String(clave || '').trim();
    const v = String(valor === null || valor === undefined ? '' : valor).trim();
    if (VALIDADORES_CONFIG[k] && !VALIDADORES_CONFIG[k](v)) return { ok: false, error: 'Valor no válido para ' + k + '.' };
    return conCandado_(() => {
      const sh = spreadsheet_().getSheetByName(SHEETS.CONFIG);
      const filas = sh.getRange(3, 1, sh.getLastRow() - 2, 2).getValues();
      const i = filas.findIndex(r => String(r[0]).trim() === k);
      if (i < 0) return { ok: false, error: 'No existe el parámetro ' + k + '.' };
      const actual = filas[i][1];
      let nuevo = v;
      if (typeof actual === 'boolean' || /^(TRUE|FALSE)$/i.test(v)) nuevo = /^(TRUE|VERDADERO|SI|SÍ|1)$/i.test(v);
      else if (v !== '' && !isNaN(Number(v)) && typeof actual === 'number') nuevo = Number(v);
      sh.getRange(3 + i, 2).setValue(nuevo);
      registrarCambio_('Config', '', k, String(actual), String(nuevo));
      invalidarConfig_();
      return { ok: true, valor: nuevo };
    });
  });
}

function guardarInhabil(fecha, descripcion, estado) {
  return envolver_(() => {
    requerir_('config');
    const k = fechaKeyDeHoja_(fecha);
    if (!k) return { ok: false, error: 'Fecha no válida.' };
    const est = String(estado).toUpperCase() === 'QUITAR' ? 'QUITAR' : 'ACTIVO';
    return conCandado_(() => {
      const sh = spreadsheet_().getSheetByName(SHEETS.INHABILES);
      const r = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(r, 1, 1, 4).setValues([[fechaDeKey_(k), descripcion || (est === 'QUITAR' ? 'Se quita de inhábiles' : 'Inhábil adicional'), 'Captura en plataforma (' + usuarioActual_().email + ')', est]]);
      sh.getRange(r, 1).setNumberFormat('dd/mm/yyyy');
      invalidarCalendario_();
      return { ok: true };
    });
  });
}

function guardarWebhookChat(url) {
  return envolver_(() => {
    requerir_('config');
    const u = String(url || '').trim();
    if (u && !/^https:\/\/chat\.googleapis\.com\//.test(u)) return { ok: false, error: 'El webhook debe empezar con https://chat.googleapis.com/' };
    if (u) PropertiesService.getScriptProperties().setProperty('CHAT_WEBHOOK_URL', u);
    else PropertiesService.getScriptProperties().deleteProperty('CHAT_WEBHOOK_URL');
    return { ok: true, configurado: !!u };
  });
}

function probarChat() {
  return envolver_(() => { requerir_('config'); return chatEnviar_('Prueba desde la plataforma de Cobranza Preventiva. Si ves este mensaje, el espacio quedó conectado.', false); });
}

function instalarTriggersUI() {
  return envolver_(() => { requerir_('config'); return instalarTriggers(); });
}

function setFechaSimulada(k) {
  return envolver_(() => {
    requerir_('config');
    if (!cfgBool_('MODO_PRUEBA', false)) return { ok: false, error: 'Sólo se puede simular una fecha con MODO_PRUEBA activo.' };
    const p = PropertiesService.getScriptProperties();
    if (!k) p.deleteProperty('FECHA_SIMULADA');
    else if (/^\d{4}-\d{2}-\d{2}$/.test(k)) p.setProperty('FECHA_SIMULADA', k);
    else return { ok: false, error: 'Usa el formato aaaa-mm-dd.' };
    return { ok: true };
  });
}

// ─── CARTERA Y FICHA DE CLIENTE ────────────────────────────────────────────

function getCartera() { return envolver_(() => armarCartera_()); }
function getFichaCliente(linea) { return envolver_(() => armarFichaCliente_(linea)); }
function calcularSaldoAFecha(linea, fecha) { return envolver_(() => calcularSaldoAFecha_(linea, String(fecha || ''))); }
function getResumenInicio() { return envolver_(() => armarResumenInicio_()); }
function getHojaTrabajo() { return envolver_(() => armarHojaTrabajo_()); }
function getSaldosVencidos() { return envolver_(() => armarSaldosVencidos_()); }
function getReporteCargado(tipo) { return envolver_(() => leerReporteCargado_(tipo)); }
function getCortes() { return envolver_(() => ({ ok: true, cortes: leerCortes_(80), catalogos: { tasas: leerTasas_().size, correos: leerContactos_().size } })); }
