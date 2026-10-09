const { makeContext, loadGs } = require('./load');
const assert = require('assert');
const ctx = makeContext({ SpreadsheetApp: { openById: () => { throw new Error('sin hoja'); } }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) } });
loadGs(ctx, ['Config.gs', 'Calendario.gs', 'Tandas.gs']);
const run = s => require('vm').runInContext(s, ctx);
let ok = 0; const t = (n, f) => { f(); ok++; console.log('ok -', n); };

t('reglas 2026 = lista CNBV 2026 (verificada en UnoTV)', () => {
  const cnbv = ['2026-01-01','2026-02-02','2026-03-16','2026-04-02','2026-04-03','2026-05-01','2026-09-16','2026-11-02','2026-11-16','2026-12-12','2026-12-25'];
  const r = run('inhabilesPorRegla_(2026).map(x=>x[0])');
  assert.deepStrictEqual([...r], cnbv);
});
t('2027 ya no cae como hábil el 1-ene', () => { assert.strictEqual(run("esHabil_('2027-01-01')"), false); });
t('Semana Santa 2027 (Pascua 28-mar)', () => { assert.strictEqual(run("pascua_(2027)"), '2027-03-28'); });
t('fin de semana y hábil', () => {
  assert.strictEqual(run("esHabil_('2026-10-10')"), false); // sábado
  assert.strictEqual(run("esHabil_('2026-10-12')"), true);  // lunes
});
t('pago efectivo: domingo 30-ago-2026 -> lunes 31', () => assert.strictEqual(run("siguienteHabil_('2026-08-30')"), '2026-08-31'));
t('sumarHabiles: vie 9-oct + 5 = vie 16-oct', () => assert.strictEqual(run("sumarHabiles_('2026-10-09',5)"), '2026-10-16'));
t('restarHabiles: lun 12-oct - 1 = vie 9-oct', () => assert.strictEqual(run("restarHabiles_('2026-10-12',1)"), '2026-10-09'));
t('habilesEntre sáb->lun = 1; vie->vie siguiente = 5', () => {
  assert.strictEqual(run("habilesEntre_('2026-10-10','2026-10-12')"), 1);
  assert.strictEqual(run("habilesEntre_('2026-10-09','2026-10-16')"), 5);
  assert.strictEqual(run("habilesEntre_('2026-10-16','2026-10-09')"), -5);
});
t('feriado intermedio: mar 15-sep-2026 + 1 hábil = jue 17 (el 16 es inhábil)', () => assert.strictEqual(run("sumarHabiles_('2026-09-15',1)"), '2026-09-17'));
t('ventana L=1, viernes 9-oct: pagos 12→16 oct, nominal sáb 10→vie 16', () => {
  const v = run("ventanaTanda_('2026-10-09',1,5,1)");
  assert.strictEqual(v.pagoDesde, '2026-10-12'); assert.strictEqual(v.pagoHasta, '2026-10-16');
  assert.strictEqual(v.nominalDesde, '2026-10-10'); assert.strictEqual(v.nominalHasta, '2026-10-16');
  assert.strictEqual(v.dias[0].aviso1Pago, '2026-10-16'); assert.strictEqual(v.dias[0].aviso2Pago, '2026-10-12');
});
t('ventana L=1, lunes 12-oct: nominal mar 13 → lun 19', () => {
  const v = run("ventanaTanda_('2026-10-12',1,5,1)");
  assert.strictEqual(v.nominalDesde, '2026-10-13'); assert.strictEqual(v.nominalHasta, '2026-10-19');
});
t('ventana con inhábil: lunes 14-sep-2026 (16-sep inhábil)', () => {
  const v = run("ventanaTanda_('2026-09-14',1,5,1)");
  assert.strictEqual(v.pagoDesde, '2026-09-15'); assert.strictEqual(v.pagoHasta, '2026-09-22'); // 15,17,18,21,22
});
t('ventana L=3 desde lunes 12-oct: pagos 13→21 (D0+1b … D0+2+5b)', () => {
  const v = run("ventanaTanda_('2026-10-12',3,5,1)");
  assert.strictEqual(v.pagoDesde, '2026-10-13'); assert.strictEqual(v.pagoHasta, '2026-10-21');
  assert.strictEqual(v.dias.length, 3); assert.strictEqual(v.dias[2].envio, '2026-10-14');
});
t('día base si hoy es sábado = lunes', () => assert.strictEqual(run("diaBaseTanda_('2026-10-10')"), '2026-10-12'));
t('estadoTanda_ con __HOY__', () => {
  ctx.__HOY__ = '2026-10-09';
  // cfg_ necesita hoja: se prueba con config vacía
  ctx.SpreadsheetApp = { openById: () => ({ getSheetByName: () => null, getSpreadsheetTimeZone: () => 'America/Mexico_City' }) };
  const e = run('estadoTanda_()');
  assert.strictEqual(e.inicioTanda, true); assert.strictEqual(e.ventana.nominalDesde, '2026-10-10');
  const marcas = e.linea.filter(x => x.aviso1 || x.aviso2).map(x => x.fecha + (x.aviso1 ? ':a1' : ':a2'));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(marcas)), ['2026-10-12:a2', '2026-10-16:a1']);
});
console.log('\n' + ok + ' pruebas OK');
