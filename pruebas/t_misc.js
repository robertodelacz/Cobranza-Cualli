const { build } = require('./mock');
const assert = require('assert'); const fs = require('fs');
let ok = 0; const t = (n, f) => { try { f(); ok++; console.log('ok -', n); } catch (e) { console.log('FALLA -', n, '\n   ', e.stack.split('\n').slice(0, 3).join('\n    ')); process.exitCode = 1; } };
const cfgSet = (m, k, v) => { const f = m.sheets.Config.grid.find(x => x[0] === k); if (f) f[1] = v; else m.sheets.Config.grid.push([k, v, '']); m.run('invalidarConfig_()'); };

// Escenario: cortes de agosto con "hoy" = 26-ago (datos reales). Corte ids vacíos → se cargan con guardarReporte para tener IDs.
function escenario(opts) {
  const m = build(Object.assign({ hoy: '2026-08-26' }, opts || {}));
  m.run('inicializarV2()');
  const d = JSON.parse(fs.readFileSync('maestro.json'));
  const iso = v => (v && v.w) ? v.w.slice(0, 10) + 'T12:00:00' : v;
  const rows1 = [['Fecha Vencimiento', 'Linea de Crédito', 'Nombre de Cliente', 'Capital', 'Interés', 'Otros', 'IVA', 'Importe', 'Moneda']].concat(d.Cache_Rep1.slice(4).filter(f => f[1]).map(f => f.map(iso)));
  const rows9 = [d.Cache_Rep9[3]].concat(d.Cache_Rep9.slice(4).filter(f => f[0]).map(f => f.map(iso)));
  m.run(`guardarReporte('rep9', ${JSON.stringify(rows9)}, {})`); m.run(`guardarReporte('rep1', ${JSON.stringify(rows1)}, {})`);
  // la carga se estampa con la fecha real del sistema; se alinea al "hoy" simulado para el escenario
  [m.sheets.Cache_Rep1, m.sheets.Cache_Rep9].forEach(sh => { sh.grid[1][1] = "'2026-08-26 11:00:00"; });
  return m;
}

// ── Víspera de verdad ──
let m = escenario();
cfgSet(m, 'MAX_EDAD_REP9_DIAS_HABILES', 10); cfgSet(m, 'MAX_EDAD_REP1_DIAS_HABILES', 10);
let c = m.plain(m.run('calcularCola_()'));
const l1 = c.items.filter(i => i.estado === 'LISTA').map(i => i.key);
m.run(`ejecutarEnvios_(${JSON.stringify(l1)}, {origen:'MANUAL'})`);
// Las bitácoras tienen sello de tiempo "real"; se reescribe a 26-ago para simular el envío de ese día
m.sheets.Bitacora_Envios.grid.slice(2).forEach(f => { f[0] = { w: '2026-08-26 10:30:00' }; });
const casos = [['2026-08-27', 'Jue 27: pago del vie 28 (dh=1) pero el aviso 1 salió ayer (<2 días hábiles) → se omite'], ['2026-08-28', 'Vie 28: pago lun 31 (dh=1), aviso 1 hace 2 días hábiles → VISPERA']];
casos.forEach(([dia, txt]) => {
  m.ctx.__HOY__ = dia; c = m.plain(m.run('calcularCola_()'));
  const dh1 = c.items.filter(i => i.diasHabiles === 1 && l1.includes(i.key));
  t(txt, () => {
    assert.ok(dh1.length > 0);
    if (dia === '2026-08-27') dh1.forEach(i => { assert.strictEqual(i.estado, 'COMPLETA'); assert.ok(/omite/.test(i.nota)); });
    else dh1.forEach(i => { assert.strictEqual(i.accion, 'VISPERA'); });
  });
});
m.ctx.__HOY__ = '2026-08-28'; c = m.plain(m.run('calcularCola_()'));
const v1 = c.items.find(i => i.accion === 'VISPERA');
t('víspera: preview dice "Recordatorio de pago" y trae el monto anterior si cambió', () => { const p = m.plain(m.run(`previsualizar_(${JSON.stringify(v1.key)})`)); assert.ok(p.ok); assert.ok(/Recordatorio de pago/.test(p.asunto)); assert.ok(p.html.includes('segundo recordatorio')); });
m.mails.length = 0;
cfgSet(m, 'MAX_EDAD_REP9_DIAS_HABILES', 10);
const sv = m.plain(m.run(`ejecutarEnvios_(${JSON.stringify([v1.key])}, {origen:'MANUAL'})`));
t('aviso 2 se envía y luego la cuota queda COMPLETA', () => { assert.strictEqual(sv.enviados, 1); const c2 = m.plain(m.run('calcularCola_()')); assert.strictEqual(c2.items.find(i => i.key === v1.key).estado, 'COMPLETA'); });

// ── Ajustes de saldo ──
m = escenario(); cfgSet(m, 'APROBADORES', 'gerente@cualli.mx');
const base = m.plain(m.run('calcularCola_()')).items.find(i => i.linea === '101306');
let a = m.plain(m.run(`crearAjuste({linea:'101306', tipo:'PAGO_APLICADO', monto:1000, motivo:'Pago recibido ayer 16:50, aún no aplicado'})`));
t('pago aplicado de $1,000 reduce el total en 1,000 y deja ajuste ACTIVO', () => {
  assert.ok(a.ok, a.error); assert.strictEqual(a.estado, 'ACTIVO');
  const i = m.plain(m.run('calcularCola_()')).items.find(x => x.linea === '101306');
  assert.ok(Math.abs(i.desglose.total - (base.desglose.total - 1000)) < 0.011); assert.ok(i.alertas.some(x => x.codigo === 'CON_AJUSTES'));
});
a = m.plain(m.run(`crearAjuste({linea:'101238', tipo:'DISPOSICION', monto:80000, motivo:'Disposición solicitada por el cliente hoy'})`));
t('ajuste grande queda PENDIENTE y bloquea la cuota', () => { assert.strictEqual(a.estado, 'PENDIENTE'); const i = m.plain(m.run('calcularCola_()')).items.find(x => x.linea === '101238'); assert.strictEqual(i.estado, 'BLOQUEADA'); assert.ok(i.bloqueos.some(b => b.codigo === 'AJUSTE_PENDIENTE')); });
t('quien captura no puede aprobar su propio ajuste', () => assert.strictEqual(m.plain(m.run(`aprobarAjuste(${JSON.stringify(a.id)})`)).ok, false));
const m2 = build({ hoy: '2026-08-26', usuario: 'gerente@cualli.mx' });
t('un ajuste exige motivo suficiente', () => assert.strictEqual(m.plain(m.run(`crearAjuste({linea:'101306', tipo:'PAGO_APLICADO', monto:10, motivo:'x'})`)).ok, false));
t('ajuste a una línea que no está en el Rep9 se rechaza', () => assert.strictEqual(m.plain(m.run(`crearAjuste({linea:'999999', tipo:'PAGO_APLICADO', monto:10, motivo:'Línea inexistente de prueba'})`)).ok, false));
// aprobar como otra persona: se reutiliza el mismo libro cambiando el usuario del mock
m.ctx.Session = { getActiveUser: () => ({ getEmail: () => 'gerente@cualli.mx' }) }; m.run('_usuario = null');
t('el aprobador de la lista puede aprobar', () => { const r = m.plain(m.run(`aprobarAjuste(${JSON.stringify(a.id)})`)); assert.ok(r.ok, r.error); const i = m.plain(m.run('calcularCola_()')).items.find(x => x.linea === '101238'); assert.notStrictEqual(i.estado, 'BLOQUEADA'); assert.ok(i.desglose.ajustes > 70000); });
// al cargar un Rep9 nuevo los ajustes dejan de aplicar
const d = JSON.parse(fs.readFileSync('maestro.json')); const iso = v => (v && v.w) ? v.w.slice(0, 10) + 'T12:00:00' : v;
m.run(`guardarReporte('rep9', ${JSON.stringify([d.Cache_Rep9[3]].concat(d.Cache_Rep9.slice(4).filter(f => f[0]).map(f => f.map(iso))))}, {})`);
t('al cargar un Rep9 nuevo, los ajustes anteriores caducan', () => { const i = m.plain(m.run('calcularCola_()')).items.find(x => x.linea === '101238'); assert.strictEqual(i.desglose.ajustes, 0); });

// ── Roles ──
m = escenario();
m.sheets.Usuarios.grid.push(['audit@cualli.mx', 'Auditoría', 'AUDITORIA', 'SI'], ['robertodelacruz@cualli.mx', 'Roberto', 'ADMIN', 'SI']);
m.ctx.Session = { getActiveUser: () => ({ getEmail: () => 'audit@cualli.mx' }) }; m.run('_usuario = null');
t('AUDITORIA puede leer pero no enviar ni cargar', () => { assert.ok(m.plain(m.run('getCola()')).ok); assert.strictEqual(m.plain(m.run(`enviarAvisos(['101306|2026-08-27'], {})`)).ok, false); assert.strictEqual(m.plain(m.run(`guardarReporte('rep1', [], {})`)).ok, false); });
m.ctx.Session = { getActiveUser: () => ({ getEmail: () => 'otro@cualli.mx' }) }; m.run('_usuario = null');
t('un correo fuera de la hoja Usuarios queda SIN_ACCESO', () => assert.strictEqual(m.plain(m.run(`enviarAvisos(['101306|2026-08-27'], {})`)).ok, false));

// ── Catálogos ──
m = escenario();
t('auditoría de catálogos detecta líneas sin tasa', () => { const a = m.plain(m.run('getCatalogos()')); assert.ok(a.ok); assert.ok(a.auditoria.sinTasa.length > 40); console.log('   sin tasa:', a.auditoria.sinTasa.length, '| sin correo:', a.auditoria.sinCorreo.length); });
t('guardar tasa (36 → 0.36) quita la advertencia SIN_TASA de 101334', () => {
  const r = m.plain(m.run(`guardarTasa('101334', 'GREEN LIGHTECH', 36)`)); assert.ok(r.ok && r.tasa === 0.36);
  const i = m.plain(m.run('calcularCola_()')).items.find(x => x.linea === '101334'); assert.ok(!i.alertas.some(a => a.codigo === 'SIN_TASA')); assert.ok(i.desglose.moratoriosProy > 0);
  assert.strictEqual(m.sheets.Cambios_Catalogo.grid.length, 2);
});
t('guardar contacto valida correos y STP de 18 dígitos', () => {
  assert.strictEqual(m.plain(m.run(`guardarContacto('101334','X','no-es-correo','646180153110000150')`)).ok, false);
  assert.strictEqual(m.plain(m.run(`guardarContacto('101334','X','a@b.com','123')`)).ok, false);
  assert.ok(m.plain(m.run(`guardarContacto('101334','X','a@b.com, c@d.mx','646180153110000150')`)).ok);
});

// ── Chat ──
m = escenario(); cfgSet(m, 'CHAT_MENCION', 'all');
t('sin webhook, el Chat falla con mensaje claro', () => assert.strictEqual(m.plain(m.run(`chatEnviar_('hola')`)).ok, false));
m.run(`guardarWebhookChat('https://chat.googleapis.com/v1/spaces/AAA/messages?key=k&token=t')`);
t('webhook sólo acepta chat.googleapis.com', () => assert.strictEqual(m.plain(m.run(`guardarWebhookChat('https://evil.example.com/x')`)).ok, false));
t('recordatorio: rango de Rep1 exacto en el mensaje', () => {
  m.run('recordatorioDescarga()'); const p = JSON.parse(m.fetches.slice(-1)[0].o.payload).text;
  assert.ok(p.includes('27/08/2026') && p.includes('02/09/2026') && p.includes('<users/all>')); console.log('   ' + p.split('\n').join('\n    '));
});
t('escalación: callada si los reportes de hoy ya están; avisa si falta el Rep1', () => {
  let n = m.fetches.length; m.run('escalarDescarga()'); assert.strictEqual(m.fetches.length, n);
  m.sheets.Cache_Rep1.grid[1][1] = "'2026-08-25 11:00:00"; m.run('escalarDescarga()'); assert.strictEqual(m.fetches.length, n + 1);
  assert.ok(JSON.parse(m.fetches.slice(-1)[0].o.payload).text.includes('Rep1'));
});
t('instalar triggers crea recordatorio y escalación (y no envío automático si está apagado)', () => { const r = m.plain(m.run('instalarTriggers()')); assert.strictEqual(r.triggers.length, 2); });
t('fin de semana: no hay recordatorio', () => { const n = m.fetches.length; m.ctx.__HOY__ = '2026-08-29'; m.run('recordatorioDescarga()'); assert.strictEqual(m.fetches.length, n); });
console.log('\n' + ok + ' pruebas OK');
