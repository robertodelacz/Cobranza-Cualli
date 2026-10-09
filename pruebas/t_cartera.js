const { build } = require('./mock');
const assert = require('assert'); const fs = require('fs');
let ok = 0; const t = (n, f) => { try { f(); ok++; console.log('ok -', n); } catch (e) { console.log('FALLA -', n, '\n   ', e.stack.split('\n').slice(0, 3).join('\n    ')); process.exitCode = 1; } };
function escenario() {
  const m = build({ hoy: '2026-08-26' }); m.run('inicializarV2()');
  const d = JSON.parse(fs.readFileSync(__dirname + '/maestro.json'));
  const iso = v => (v && v.w) ? v.w.slice(0, 10) + 'T12:00:00' : v;
  const rows1 = [['Fecha Vencimiento', 'Linea de Crédito', 'Nombre de Cliente', 'Capital', 'Interés', 'Otros', 'IVA', 'Importe', 'Moneda']].concat(d.Cache_Rep1.slice(4).filter(f => f[1]).map(f => f.map(iso)));
  const rows9 = [d.Cache_Rep9[3]].concat(d.Cache_Rep9.slice(4).filter(f => f[0]).map(f => f.map(iso)));
  m.run(`guardarReporte('rep9', ${JSON.stringify(rows9)}, {})`); m.run(`guardarReporte('rep1', ${JSON.stringify(rows1)}, {})`);
  [m.sheets.Cache_Rep1, m.sheets.Cache_Rep9].forEach(sh => { sh.grid[1][1] = "'2026-08-26 11:00:00"; });
  ['MAX_EDAD_REP9_DIAS_HABILES', 'MAX_EDAD_REP1_DIAS_HABILES'].forEach(k => { const f = m.sheets.Config.grid.find(x => x[0] === k); f[1] = 10; }); m.run('invalidarConfig_()');
  return { m, d, rows9 };
}
const { m, rows9 } = escenario();
const c = m.plain(m.run('getCartera()'));
t('cartera: ok y trae las líneas del Rep9', () => { assert.ok(c.ok, c.error); assert.ok(c.lineas.length > 100); assert.ok(c.totales.MXN.saldoTotal > 1e6); assert.ok(c.conAvisos); });
t('cartera: orden por saldo vencido descendente', () => { for (let i = 1; i < c.lineas.length; i++) assert.ok(c.lineas[i - 1].saldoVencido >= c.lineas[i].saldoVencido - 1e-9); });
t('cartera: suma de saldos por moneda coincide con el Rep9', () => {
  const suma = rows9.slice(1).filter(r => String(r[3]).trim().toUpperCase() === 'MXN').reduce((s, r) => s + (Number(r[27]) || 0), 0);
  assert.ok(Math.abs(suma - c.totales.MXN.saldoTotal) < 0.05, suma + ' vs ' + c.totales.MXN.saldoTotal);
});
t('cartera: línea con cuota trae próxima cuota', () => { const l = c.lineas.find(x => x.linea === '101306'); assert.ok(l && l.proxima && l.proxima.total > 0); });
const f = m.plain(m.run(`getFichaCliente('101306')`));
t('ficha: saldos, cuotas y datos', () => { assert.ok(f.ok, f.error); assert.ok(f.saldos.saldoTotal > 0); assert.ok(f.cuotas.length >= 1); assert.strictEqual(f.linea, '101306'); assert.ok('correo' in f.contacto); });
t('ficha: línea inexistente da error claro', () => { const r = m.plain(m.run(`getFichaCliente('999999')`)); assert.strictEqual(r.ok, false); });
// el cálculo a la fecha de la cuota coincide con el total del aviso (misma fórmula, base nominal)
const cola = m.plain(m.run('calcularCola_()'));
const conVenc = cola.items.filter(i => i.desglose.capVencido > 0 && i.tasaMoratoria > 0 && !i.multiCuota);
t('hay casos con capital vencido y tasa para comparar', () => assert.ok(conVenc.length > 0));
conVenc.slice(0, 3).forEach(i => t('saldo a fecha = total del aviso (' + i.linea + ')', () => {
  const r = m.plain(m.run(`calcularSaldoAFecha('${i.linea}', '${i.fechaNominal}')`));
  assert.ok(r.ok, r.error); assert.ok(Math.abs(r.desglose.total - i.total) < 0.011, r.desglose.total + ' vs ' + i.total);
}));
t('saldo a fecha: fecha anterior al corte se rechaza', () => assert.strictEqual(m.plain(m.run(`calcularSaldoAFecha('101306','2026-08-01')`)).ok, false));
t('saldo a fecha: fecha inválida se rechaza', () => assert.strictEqual(m.plain(m.run(`calcularSaldoAFecha('101306','mañana')`)).ok, false));
t('saldo a fecha: más lejos, más moratorios (monotonía)', () => {
  const i = conVenc[0]; const a = m.plain(m.run(`calcularSaldoAFecha('${i.linea}','${i.fechaNominal}')`)); const b = m.plain(m.run(`calcularSaldoAFecha('${i.linea}','2026-09-30')`));
  assert.ok(b.desglose.moratoriosProy > a.desglose.moratoriosProy);
});
const r = m.plain(m.run('getResumenInicio()'));
t('resumen de inicio', () => { assert.ok(r.ok && r.hayRep9); assert.ok(r.cartera.MXN.lineas > 100); assert.strictEqual(typeof r.avisosSemana, 'number'); });
t('getBitacora sigue funcionando tras el refactor', () => { const b = m.plain(m.run('getBitacora({limite:50})')); assert.ok(b.ok && Array.isArray(b.registros)); });
const vacio = build({ hoy: '2026-08-26' }); vacio.run('inicializarV2()'); ['Cache_Rep1', 'Cache_Rep9'].forEach(n => { vacio.sheets[n].grid = vacio.sheets[n].grid.slice(0, 3); });
t('sin Rep9: cartera avisa con claridad', () => { const x = vacio.plain(vacio.run('getCartera()')); assert.strictEqual(x.ok, false); assert.ok(/Rep9/.test(x.error)); });
console.log(ok + ' pruebas');
