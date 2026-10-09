const { build } = require('./mock');
const assert = require('assert'); const fs = require('fs');
let ok = 0; const t = (n, f) => { try { f(); ok++; console.log('ok -', n); } catch (e) { console.log('FALLA -', n, '\n   ', e.stack.split('\n').slice(0, 3).join('\n    ')); process.exitCode = 1; } };

const m = build({ hoy: '' });
const HOY = m.run("Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')");
m.ctx.__HOY__ = HOY;
console.log('Hoy del sistema:', HOY, '| tz inicial de la hoja:', m.ss.getSpreadsheetTimeZone());

// vaciar cachés y bitácora de prueba para simular instalación limpia del flujo v2
['Cache_Rep1', 'Cache_Rep9'].forEach(n => { const g = m.sheets[n].grid; m.sheets[n].grid = g.slice(0, 4); });
m.sheets.Bitacora_Envios.grid = m.sheets.Bitacora_Envios.grid.slice(0, 2);

// 1. Inicialización
let r = m.plain(m.run('inicializarV2()'));
t('inicializarV2 crea hojas, parámetros y cambia la zona horaria', () => {
  assert.ok(r.ok); assert.strictEqual(m.ss.getSpreadsheetTimeZone(), 'America/Mexico_City');
  ['Cortes', 'Ajustes', 'Calendario_Inhabiles', 'Usuarios', 'Cambios_Catalogo'].forEach(n => assert.ok(m.sheets[n], n));
  const claves = m.sheets.Config.grid.map(f => f[0]);
  assert.ok(claves.includes('TANDA_DIAS_HABILES') && claves.includes('BASE_MORATORIOS'));
  assert.strictEqual(m.sheets.Bitacora_Envios.grid[1].length >= 20, true);
});
r = m.plain(m.run('inicializarV2()'));
t('inicializarV2 es idempotente (no duplica parámetros)', () => { const claves = m.sheets.Config.grid.map(f => f[0]).filter(Boolean); assert.strictEqual(claves.length, new Set(claves).size); });
t('calendario visible: 33 fechas 2026-2028', () => assert.strictEqual(m.sheets.Calendario_Inhabiles.grid.length - 1, 33));

// 2. Ventana y reportes sintéticos a partir de los datos reales
const e = m.plain(m.run('estadoTanda_()'));
console.log('Ventana de hoy:', e.ventana.nominalDesde, '→', e.ventana.nominalHasta, '| inicio de tanda:', e.inicioTanda);
const iso = k => k + 'T12:00:00';
const fechas = []; for (let k = e.ventana.nominalDesde; k <= e.ventana.nominalHasta; k = m.run(`addDays_('${k}',1)`)) fechas.push(k);
const real1 = JSON.parse(fs.readFileSync('maestro.json')).Cache_Rep1.slice(4).filter(f => f[1]);
const rows1 = [['Reporte de vencimientos'], ['Fecha Vencimiento', 'Linea de Crédito', 'Nombre de Cliente', 'Capital', 'Interés', 'Otros', 'IVA', 'Importe', 'Moneda']]
  .concat(real1.map((f, i) => [iso(fechas[i % fechas.length]), f[1], f[2], f[3], f[4], f[5], f[6], f[7], f[8]]));
const g9 = JSON.parse(fs.readFileSync('maestro.json')).Cache_Rep9;
const toIso = v => (v && v.w) ? v.w.slice(0, 10) + 'T12:00:00' : v;
const rows9 = [g9[3]].concat(g9.slice(4).filter(f => f[0]).map(f => f.map(toIso)));

let v = m.plain(m.run(`validarReporte('rep1', ${JSON.stringify(rows1)}, {})`));
t('validar Rep1: cobertura calculada y sin errores', () => { assert.ok(v.ok, v.error); assert.strictEqual(v.filas, 19); assert.strictEqual(v.requerido.desde, e.ventana.nominalDesde); });
v = m.plain(m.run(`validarReporte('rep9', ${JSON.stringify(rows9)}, {})`));
t('validar Rep9: avisa que sólo trae MXN', () => { assert.ok(v.ok, v.error); assert.ok(v.avisos.some(a => /sólo trae líneas en MXN/.test(a.texto))); });
t('Rep9 con encabezado equivocado se rechaza', () => { const bad = JSON.parse(JSON.stringify(rows9)); bad[0][5] = 'Otra'; const x = m.plain(m.run(`validarReporte('rep9', ${JSON.stringify(bad)}, {})`)); assert.strictEqual(x.ok, false); });
let g = m.plain(m.run(`guardarReporte('rep9', ${JSON.stringify(rows9)}, {archivo:'Rep9.xlsx'})`));
t('guardar Rep9 → corte registrado', () => { assert.ok(g.ok, g.error); assert.ok(/^R9-/.test(g.corteId)); });
g = m.plain(m.run(`guardarReporte('rep1', ${JSON.stringify(rows1)}, {archivo:'Rep1.xlsx', rangoDesde:'${e.ventana.nominalDesde}', rangoHasta:'${e.ventana.nominalHasta}'})`));
t('guardar Rep1 → corte registrado con rango declarado', () => { assert.ok(g.ok, g.error); assert.strictEqual(m.sheets.Cortes.grid.length, 3); assert.ok(m.sheets.Cortes.grid[2][8].includes(' a ')); });

// 3. Cola
let c = m.plain(m.run('calcularCola_()'));
t('cola con cortes de hoy: hay cuotas por enviar', () => { assert.ok(c.ok, c.error); assert.ok(c.stats.porEnviar > 0); console.log('   ', JSON.stringify(c.stats.porEstado)); });
t('ninguna cuota queda como vencida por fecha mal leída (fechas con tz de hoja cambiada)', () => assert.strictEqual(c.items.filter(i => i.estado === 'VENCIDA').length, 0));

// 4. Envío en MODO_PRUEBA: no consume el aviso real
m.sheets.Config.grid.find(f => f[0] === 'MODO_PRUEBA')[1] = true; m.run('invalidarConfig_()');
const listas = c.items.filter(i => i.estado === 'LISTA').map(i => i.key);
let s = m.plain(m.run(`ejecutarEnvios_(${JSON.stringify(listas.slice(0, 5))}, {origen:'MANUAL'})`));
t('modo prueba: se envían a buzón interno con [PRUEBA]', () => { assert.ok(s.ok, s.error); assert.strictEqual(s.enviados, 5); assert.ok(m.mails.every(x => x.to === 'robertodelacruz@cualli.mx' && x.subj.length > 0)); });
c = m.plain(m.run('calcularCola_()'));
t('modo prueba NO consume el aviso (siguen LISTA)', () => assert.ok(listas.slice(0, 5).every(k => c.items.find(i => i.key === k).estado === 'LISTA')));
const filaPrueba = m.sheets.Bitacora_Envios.grid[2];
t('bitácora: estado ENVIADO_PRUEBA y fecha nominal en columna B', () => { assert.strictEqual(filaPrueba[7], 'ENVIADO_PRUEBA'); assert.ok(filaPrueba[1].w.endsWith('12:00:00')); });

// 5. Envío real
m.mails.length = 0;
m.sheets.Config.grid.find(f => f[0] === 'MODO_PRUEBA')[1] = false; m.run('invalidarConfig_()');
s = m.plain(m.run(`ejecutarEnvios_(${JSON.stringify(listas)}, {origen:'MANUAL'})`));
t('envío real: todas las LISTAS salen', () => { assert.ok(s.ok); assert.strictEqual(s.enviados, listas.length); assert.strictEqual(s.errores, 0); });
t('envío real: reply-to presente y SIN noReply; asunto codificado; HTML con aviso legal', () => {
  const x = m.mails[0]; assert.ok(x.adv.replyTo && x.adv.noReply === undefined); assert.ok(/^=\?UTF-8\?B\?/.test(x.subj));
  assert.ok(x.adv.htmlBody.includes('Secretaría de Hacienda')); assert.ok(!/PRUEBA/.test(Buffer.from(x.subj.slice(10, -2), 'base64').toString()));
});
c = m.plain(m.run('calcularCola_()'));
t('después de enviar: pasan a ESPERA_VISPERA o COMPLETA (nunca LISTA otra vez)', () => assert.ok(listas.every(k => ['ESPERA_VISPERA', 'COMPLETA'].includes(c.items.find(i => i.key === k).estado))));
m.mails.length = 0;
s = m.plain(m.run(`ejecutarEnvios_(${JSON.stringify(listas.slice(0, 3))}, {origen:'MANUAL'})`));
t('doble clic / reintento: no duplica (0 enviados, omitidos con motivo)', () => { assert.strictEqual(s.enviados, 0); assert.strictEqual(s.omitidos, 3); assert.strictEqual(m.mails.length, 0); });

// 6. REVISAR requiere confirmación; reenvío; fallo de un envío no tumba el lote
const revisar = c.items.filter(i => i.estado === 'REVISAR').map(i => i.key);
if (revisar.length) {
  s = m.plain(m.run(`ejecutarEnvios_(${JSON.stringify(revisar.slice(0, 1))}, {origen:'MANUAL'})`));
  t('REVISAR sin confirmación se omite', () => { assert.strictEqual(s.enviados, 0); assert.ok(/alertas/.test(s.resultados[0].error)); });
  s = m.plain(m.run(`ejecutarEnvios_(${JSON.stringify(revisar.slice(0, 1))}, {origen:'MANUAL', confirmarRevision:true})`));
  t('REVISAR con confirmación sale', () => assert.strictEqual(s.enviados, 1));
}
m.mails.length = 0;
const reenv = m.plain(m.run(`reenviarAviso(${JSON.stringify(listas[0])}, 'El cliente pidió que se reenvíe')`));
t('reenvío con motivo: sale y queda como REENVIO en bitácora', () => { assert.ok(reenv.ok && reenv.enviados === 1, reenv.error); const f = m.sheets.Bitacora_Envios.grid.slice(-1)[0]; assert.strictEqual(f[2], 'REENVIO'); assert.strictEqual(f[17], 'REENVIO'); });
t('reenvío sin motivo se rechaza', () => assert.strictEqual(m.plain(m.run(`reenviarAviso(${JSON.stringify(listas[0])}, '')`)).ok, false));
c = m.plain(m.run('calcularCola_()'));
t('reenvío no altera el calendario de avisos', () => assert.ok(['ESPERA_VISPERA', 'COMPLETA'].includes(c.items.find(i => i.key === listas[0]).estado)));

// 7. Aviso 2 (víspera) cuando llega el día: se simula avanzar la fecha
const k1 = c.items.filter(i => i.estado === 'ESPERA_VISPERA')[0];
if (k1) {
  m.sheets.Config.grid.find(f => f[0] === 'MAX_EDAD_REP9_DIAS_HABILES')[1] = 10; m.sheets.Config.grid.find(f => f[0] === 'TANDA_DIAS_HABILES')[1] = 10; m.run('invalidarConfig_()');
  // el día previo a 1 día hábil del pago
  const dv = m.run(`restarHabiles_('${k1.fechaPago}', 1)`);
  m.ctx.__HOY__ = dv; m.run('invalidarConfig_()');
  c = m.plain(m.run('calcularCola_()'));
  const it = c.items.find(i => i.key === k1.key);
  t('víspera: en el día hábil anterior al pago le toca aviso 2 (' + dv + ')', () => { assert.strictEqual(it.diasHabiles, 1); assert.strictEqual(it.accion, it.previo && m.run(`habilesEntre_('${it.previo.fechaKey}','${dv}')`) >= 2 ? 'VISPERA' : null); });
}
console.log('\n' + ok + ' pruebas OK');
fs.writeFileSync('mails_flow.json', JSON.stringify(m.mails.slice(0, 2)));
