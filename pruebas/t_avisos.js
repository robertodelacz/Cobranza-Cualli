const { build } = require('./mock');
const assert = require('assert');
let ok = 0; const t = (n, f) => { try { f(); ok++; console.log('ok -', n); } catch (e) { console.log('FALLA -', n, '\n   ', e.message); process.exitCode = 1; } };
const m = build({ hoy: '2026-10-23' });
const cal = (pago, hoy, reg, vig) => m.plain(m.run('calendarioAvisos_(' + JSON.stringify(pago) + ', habilesEntre_(' + JSON.stringify(hoy) + ',' + JSON.stringify(pago) + '), ' + JSON.stringify(reg || { previo: null, vispera: null }) + ', diasAviso_(), ' + JSON.stringify(hoy) + ', ' + JSON.stringify(vig || null) + ')'));

t('vie 23-oct, paga mar 27: el primero ya iba tarde (sale hoy) y el segundo se omite', () => {
  const c = cal('2026-10-27', '2026-10-23');
  assert.strictEqual(c.a1.estado, 'ATRASADO'); assert.strictEqual(c.a1.fecha, '2026-10-23'); assert.strictEqual(c.a1.previsto, '2026-10-20');
  assert.strictEqual(c.a2.estado, 'OMITIDO'); assert.strictEqual(c.a2.fecha, '2026-10-26');
});
t('paga mié 4-nov: primero el 27-oct y segundo el 3-nov (el 2-nov es inhábil)', () => {
  const c = cal('2026-11-04', '2026-10-23');
  assert.strictEqual(c.a1.estado, 'PROGRAMADO'); assert.strictEqual(c.a1.fecha, '2026-10-27');
  assert.strictEqual(c.a2.estado, 'PROGRAMADO'); assert.strictEqual(c.a2.fecha, '2026-11-03');
});
t('paga vie 30-oct: el primero es hoy (5 hábiles antes) y el segundo el jue 29', () => {
  const c = cal('2026-10-30', '2026-10-23');
  assert.strictEqual(c.a1.estado, 'HOY'); assert.strictEqual(c.a2.fecha, '2026-10-29'); assert.strictEqual(c.a2.estado, 'PROGRAMADO');
});
t('primer aviso ya enviado: muestra la fecha real y el segundo sigue su plan', () => {
  const c = cal('2026-10-30', '2026-10-26', { previo: { fechaKey: '2026-10-23' }, vispera: null });
  assert.strictEqual(c.a1.estado, 'ENVIADO'); assert.strictEqual(c.a1.fecha, '2026-10-23');
  assert.strictEqual(c.a2.fecha, '2026-10-29'); assert.strictEqual(c.a2.estado, 'PROGRAMADO');
});
t('ambos enviados: los dos aparecen como enviados', () => {
  const c = cal('2026-10-30', '2026-10-29', { previo: { fechaKey: '2026-10-23' }, vispera: { fechaKey: '2026-10-29' } });
  assert.strictEqual(c.a1.estado, 'ENVIADO'); assert.strictEqual(c.a2.estado, 'ENVIADO'); assert.strictEqual(c.a2.fecha, '2026-10-29');
});
t('vencida sin avisos: no hay fechas', () => { const c = cal('2026-10-20', '2026-10-23'); assert.strictEqual(c.a1, null); assert.strictEqual(c.a2, null); });
t('fecha fuera de la vigencia de los reportes → cubierto=false', () => {
  const c = cal('2026-11-04', '2026-10-23', null, '2026-10-29');
  assert.strictEqual(c.a1.cubierto, true); assert.strictEqual(c.a2.cubierto, false);
});
t('vigencia: Rep1 y Rep9 del 22-oct con ciclo 5 y Rep9 máx 4 valen hasta el 28-oct (envíos 22 a 28)', () => {
  const m2 = build({ hoy: '2026-10-23' });
  const cfg = (k, v) => { const f = m2.sheets.Config.grid.find(x => x[0] === k); if (f) f[1] = v; else m2.sheets.Config.grid.push([k, v, '']); m2.run('invalidarConfig_()'); };
  cfg('TANDA_DIAS_HABILES', 5); cfg('MAX_EDAD_REP9_DIAS_HABILES', 4); cfg('MAX_EDAD_REP1_DIAS_HABILES', '');
  const v = m2.run("vigenciaReportes_({ rep1: { corte: { fechaKey: '2026-10-22' }, max: 4 }, rep9: { corte: { fechaKey: '2026-10-22' }, max: 4 } })");
  assert.strictEqual(v, '2026-10-28');
});
t('el motor trae calendario en cada cuota y la vigencia en el resultado', () => {
  const m3 = build({ hoy: '2026-08-26' });
  const c = m3.plain(m3.run('calcularCola_()'));
  assert.ok(c.items.every(i => 'calendario' in i)); assert.ok('vigenciaHasta' in c);
  assert.ok(c.items.filter(i => i.accion === 'PREVENTIVO').every(i => i.calendario.a1 && (i.calendario.a1.estado === 'HOY' || i.calendario.a1.estado === 'ATRASADO')));
});
console.log('\n' + ok + ' pruebas OK');
