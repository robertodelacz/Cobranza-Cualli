const { build } = require('./mock');
const assert = require('assert');
let ok = 0; const t = (n, f) => { try { f(); ok++; console.log('ok -', n); } catch (e) { console.log('FALLA -', n, '\n   ', e.message); process.exitCode = 1; } };

// ── 1. Motor sobre los datos reales (como estaban el 26-ago-2026) ──
let m = build({ hoy: '2026-08-26' });
let c = m.plain(m.run('calcularCola_()'));
t('motor corre sobre el Excel real', () => assert.strictEqual(c.ok, true, c.error));
t('19 cuotas del Rep1', () => assert.strictEqual(c.items.length, 19));
t('3 vencen hoy (26-ago) y no reciben aviso', () => assert.strictEqual(c.items.filter(i => i.estado === 'VENCE_HOY').length, 3));
t('16 por enviar: aviso 1 (todas son primera vez)', () => { assert.strictEqual(c.stats.porEnviar, 16); assert.strictEqual(c.stats.aviso1, 16); });
t('sáb 29 y dom 30 pagan lunes 31; días hábiles = 3', () => {
  const fin = c.items.filter(i => i.fechaNominal === '2026-08-29' || i.fechaNominal === '2026-08-30');
  assert.strictEqual(fin.length, 9);
  fin.forEach(i => { assert.strictEqual(i.fechaPago, '2026-08-31'); assert.strictEqual(i.diasHabiles, 3); });
});
t('Rep1 y Rep9 del mismo día: sin bloqueos de frescura', () => assert.ok(!c.items.some(i => i.bloqueos.some(b => /VENCIDO/.test(b.codigo)))));
const por = id => c.items.find(i => i.linea === id);
t('101334: sin tasa + capital vencido → REVISAR y moratorios proyectados = 0', () => {
  const i = por('101334'); assert.strictEqual(i.estado, 'REVISAR'); assert.strictEqual(i.desglose.moratoriosProy, 0);
  assert.ok(i.alertas.some(a => a.codigo === 'SIN_TASA' && a.nivel === 'warn'));
});
t('101335 (USD): sin Rep9 → REVISAR, moneda USD', () => { const i = por('101335'); assert.strictEqual(i.estado, 'REVISAR'); assert.strictEqual(i.moneda, 'USD'); assert.ok(i.sinRep9); });
t('101238: con tasa y capital vencido → moratorios proyectados desde el corte del Rep9', () => {
  const i = por('101238'); assert.ok(i.desglose.capVencido > 50000); assert.strictEqual(i.diasProy, 1);
  assert.ok(Math.abs(i.desglose.moratoriosProy - i.desglose.capVencido * i.tasaMoratoria / 360 * 1) < 0.01);
});
t('total = cuota + vencido + moratorios (+ajustes)', () => c.items.forEach(i => { const d = i.desglose; assert.ok(Math.abs(d.total - (d.cuota + d.capVencido + d.intVencidos + d.moratoriosAcum + d.moratoriosProy + d.ajustes)) < 0.011); }));
t('Importe del Rep1 cuadra con sus componentes en las 19', () => assert.ok(!c.items.some(i => i.alertas.some(a => a.codigo === 'IMPORTE_NO_CUADRA'))));
t('las demás con tasa y sin vencido están LISTAS', () => assert.ok(c.stats.porEstado.LISTA >= 9));
console.log('   estados:', JSON.stringify(c.stats.porEstado), '| MXN', c.stats.sumaMXN, '| USD', c.stats.sumaUSD);

// ── 2. Reportes viejos: el 9-oct el caché de agosto debe BLOQUEAR el envío ──
m = build({ hoy: '2026-10-09' });
c = m.plain(m.run('calcularCola_()'));
t('el 9-oct con reportes de agosto: todas vencidas (no hay nada que avisar)', () => assert.ok(c.items.every(i => i.estado === 'VENCIDA')));
m = build({ hoy: '2026-08-31' }); c = m.plain(m.run('calcularCola_()'));
t('31-ago con reportes del 26: Rep1/Rep9 vencidos bloquean (antes se enviaba con montos viejos)', () => {
  const el = c.items.filter(i => i.accion);
  assert.ok(c.frescura.rep9.ok === false);
  el.forEach(i => assert.ok(i.bloqueos.some(b => b.codigo === 'REP9_VENCIDO')));
  assert.ok(el.every(i => i.estado === 'BLOQUEADA'));
});

// ── 3. Ventana de tanda para la fecha de los datos ──
m = build({ hoy: '2026-08-26' });
let e = m.plain(m.run('estadoTanda_()'));
t('26-ago: Rep1 debía pedirse del 27-ago al 2-sep (el cargado era 26→30)', () => { assert.strictEqual(e.ventana.nominalDesde, '2026-08-27'); assert.strictEqual(e.ventana.nominalHasta, '2026-09-02'); });
console.log('\n' + ok + ' pruebas OK');
