const { build } = require('./mock');
const assert = require('assert'); const fs = require('fs');
let ok = 0; const t = (n, f) => { try { f(); ok++; console.log('ok -', n); } catch (e) { console.log('FALLA -', n, '\n   ', e.stack.split('\n').slice(0, 3).join('\n    ')); process.exitCode = 1; } };
const m = build({ hoy: '2026-08-26' }); m.run('inicializarV2()');
const d = JSON.parse(fs.readFileSync(__dirname + '/maestro.json'));
const iso = v => (v && v.w) ? v.w.slice(0, 10) + 'T12:00:00' : v;
const rows1 = [['Fecha Vencimiento', 'Linea de Crédito', 'Nombre de Cliente', 'Capital', 'Interés', 'Otros', 'IVA', 'Importe', 'Moneda']].concat(d.Cache_Rep1.slice(4).filter(f => f[1]).map(f => f.map(iso)));
const rows9 = [d.Cache_Rep9[3]].concat(d.Cache_Rep9.slice(4).filter(f => f[0]).map(f => f.map(iso)));
['MAX_EDAD_REP9_DIAS_HABILES', 'MAX_EDAD_REP1_DIAS_HABILES'].forEach(k => { const f = m.sheets.Config.grid.find(x => x[0] === k); f[1] = 10; }); m.run('invalidarConfig_()');
const g9 = m.plain(m.run(`guardarReporte('rep9', ${JSON.stringify(rows9)}, {archivo:'Rep9.xlsx'})`));
const g1 = m.plain(m.run(`guardarReporte('rep1', ${JSON.stringify(rows1)}, {archivo:'Rep1.xlsx'})`));
[m.sheets.Cache_Rep1, m.sheets.Cache_Rep9].forEach(sh => { sh.grid[1][1] = "'2026-08-26 11:00:00"; });

const cola = m.plain(m.run('calcularCola_()'));
const h = m.plain(m.run('getHojaTrabajo()'));
t('cada cuota trae su rastro (traza)', () => { assert.ok(cola.items.length > 5); cola.items.forEach(i => { assert.ok(i.traza && i.traza.rep1); assert.strictEqual(i.traza.baseDias, 360); }); });
t('traza: lo vencido del Rep9 coincide con el desglose', () => {
  cola.items.filter(i => i.traza.aplicaVencido).forEach(i => {
    const r = i.traza.rep9;
    assert.ok(Math.abs(r.capVencido - i.desglose.capVencido) < 0.011);
    assert.ok(Math.abs(r.intVencido + r.ivaIntVencido - i.desglose.intVencidos) < 0.011);
    assert.ok(Math.abs(r.moratorios + r.ivaMoratorios + r.morCont + r.ivaMorCont - i.desglose.moratoriosAcum) < 0.011);
  });
});
t('traza: moratorios del periodo = capVencido × tasa moratoria ÷ 360 × días', () => {
  const con = cola.items.filter(i => i.desglose.moratoriosProy > 0);
  assert.ok(con.length > 0);
  con.forEach(i => { const esp = i.desglose.capVencido * i.tasaMoratoria / 360 * i.diasProy; assert.ok(Math.abs(esp - i.desglose.moratoriosProy) < 0.011, i.key); });
});
t('hoja de trabajo: ok, mismas filas que los avisos', () => { assert.ok(h.ok, h.error); assert.strictEqual(h.filas.length, cola.items.length); assert.strictEqual(h.base, 'NOMINAL'); assert.strictEqual(h.factor, 2); });
t('hoja de trabajo: TOTAL = D+E+F+G+J+K+L+M+ajustes en cada fila', () => {
  h.filas.forEach(f => { const s = f.capital + f.intereses + f.otros + f.iva + f.capVencido + f.intVencidos + f.sumaMoratorios + f.moratoriosPeriodo + f.ajustes; assert.ok(Math.abs(s - f.total) < 0.021, f.key + ' ' + s + ' vs ' + f.total); });
});
t('hoja de trabajo: totales por moneda cuadran con la suma de filas', () => {
  Object.keys(h.totales).forEach(mon => { const s = h.filas.filter(f => f.moneda === mon).reduce((a, f) => a + f.total, 0); assert.ok(Math.abs(s - h.totales[mon].total) < 0.05); });
});
const sv = m.plain(m.run('getSaldosVencidos()'));
t('saldos vencidos: suma moratorios = R+S+T+U e intereses vencidos = P+Q', () => {
  assert.ok(sv.ok, sv.error); assert.ok(sv.filas.length > 100);
  sv.filas.forEach(f => { assert.ok(Math.abs(f.sumaMoratorios - (f.moratorios + f.ivaMoratorios + f.morCont + f.ivaMorCont)) < 0.021); assert.ok(Math.abs(f.interesesVencidos - (f.intVencido + f.ivaIntVencido)) < 0.011); });
});
t('saldos vencidos: saldo total por moneda coincide con el Rep9', () => {
  const suma = rows9.slice(1).filter(r => String(r[3]).trim().toUpperCase() === 'MXN').reduce((s, r) => s + (Number(r[27]) || 0), 0);
  assert.ok(Math.abs(suma - sv.totales.MXN.saldoTotal) < 0.05);
});
const r1 = m.plain(m.run(`getReporteCargado('rep1')`)), r9 = m.plain(m.run(`getReporteCargado('rep9')`));
t('reporte cargado: Rep1 trae encabezados y filas con fechas como texto', () => { assert.ok(r1.ok, r1.error); assert.strictEqual(r1.encabezados.length, 9); assert.ok(r1.filas.length > 5); assert.ok(/^\d{4}-\d{2}-\d{2}/.test(String(r1.filas[0][0])), JSON.stringify(r1.filas[0])); });
t('reporte cargado: Rep9 trae 30 columnas', () => { assert.ok(r9.ok); assert.strictEqual(r9.encabezados.length, 30); assert.strictEqual(r9.filas[0].length, 30); });
t('reporte cargado: tipo inválido da error', () => assert.strictEqual(m.plain(m.run(`getReporteCargado('x')`)).ok, false));
const ct = m.plain(m.run('getCortes()'));
t('cortes: quedan las dos cargas, la más reciente primero', () => { assert.ok(ct.ok, ct.error); assert.ok(ct.cortes.length >= 2); assert.ok(ct.cortes.some(c => c.reporte === 'Rep1' && c.archivo === 'Rep1.xlsx')); assert.ok(ct.cortes.some(c => c.reporte === 'Rep9' && c.filas > 100)); });
console.log(ok + ' pruebas');
