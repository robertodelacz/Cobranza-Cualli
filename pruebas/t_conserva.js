const { build } = require('./mock');
const assert = require('assert');
let ok = 0; const t = (n, f) => { try { f(); ok++; console.log('ok -', n); } catch (e) { console.log('FALLA -', n, '\n   ', e.message); process.exitCode = 1; } };
const H = ['Fecha Vencimiento', 'Línea de Crédito', 'Nombre de Cliente', 'Capital', 'Interés', 'Otros', 'IVA', 'Importe', 'Moneda'];
const fila = (f, l) => [f + 'T00:00:00.000Z', l, 'CLIENTE ' + l, 1000, 100, 0, 16, 1116, 'MXN'];
const m = build({ hoy: '2026-10-23' });
const cache = () => m.plain(m.run("leerCache_(spreadsheet_().getSheetByName(SHEETS.CACHE_REP1), 9).map(r => [fechaKeyDeHoja_(r[0]), String(r[1])])"));
const cargar = filas => m.plain(m.run('guardarCache_("rep1", ' + JSON.stringify([H].concat(filas)) + ', {})'));

t('primera carga: solo 27-oct; no se conserva nada del caché de agosto (ya pasó)', () => {
  const r = cargar([fila('2026-10-27', '1001'), fila('2026-10-27', '1002'), fila('2026-10-27', '1003')]);
  assert.ok(r.ok); assert.strictEqual(r.conservadas, 0); assert.strictEqual(cache().length, 3);
});
t('cargar solo el 28: las cuotas del 27 se conservan', () => {
  const r = cargar([fila('2026-10-28', '1004'), fila('2026-10-28', '1005')]);
  assert.ok(r.ok); assert.strictEqual(r.conservadas, 3); assert.strictEqual(r.filas, 2);
  assert.strictEqual(r.conservadasDesde, '2026-10-27'); assert.strictEqual(r.conservadasHasta, '2026-10-27');
  const c = cache(); assert.strictEqual(c.length, 5);
  assert.deepStrictEqual(c.map(x => x[0]), ['2026-10-27', '2026-10-27', '2026-10-27', '2026-10-28', '2026-10-28']);
});
t('el motor calcula las 5 cuotas (27 y 28) y ninguna se duplica', () => {
  const c = m.plain(m.run('calcularCola_()'));
  assert.strictEqual(c.items.length, 5);
  assert.strictEqual(new Set(c.items.map(i => i.key)).size, 5);
});
t('recargar un rango que incluye el 27 reemplaza lo anterior: la que ya no viene se pagó', () => {
  const r = cargar([fila('2026-10-27', '1001'), fila('2026-10-28', '1004'), fila('2026-10-28', '1005')]);
  assert.strictEqual(r.conservadas, 0);
  assert.deepStrictEqual(cache().map(x => x[1]), ['1001', '1004', '1005']);
});
t('una cuota cuyo pago ya llegó no se conserva (hoy 28-oct, archivo solo del 29)', () => {
  m.ctx.__HOY__ = '2026-10-28';
  const r = cargar([fila('2026-10-29', '1006')]);
  assert.strictEqual(r.conservadas, 0);   // el 27 ya pasó y el 28 vence hoy
  assert.deepStrictEqual(cache().map(x => x[1]), ['1006']);
});
t('si el archivo trae cuotas de varios días, se conserva solo lo que queda fuera de sus fechas', () => {
  m.ctx.__HOY__ = '2026-10-23';
  cargar([fila('2026-10-27', '2001'), fila('2026-10-28', '2002'), fila('2026-11-03', '2003')]);
  const r = cargar([fila('2026-10-28', '2002'), fila('2026-10-30', '2004')]);
  assert.strictEqual(r.conservadas, 2);   // 27-oct y 3-nov quedan fuera de 28→30
  assert.deepStrictEqual(cache().map(x => x[1]), ['2001', '2002', '2004', '2003']);
});
console.log('\n' + ok + ' pruebas OK');
