// Prueba de la interfaz en un navegador real (Chromium) con el motor verdadero detrás.
// google.script.run se simula y cada llamada se ejecuta contra los .gs cargados en Node.
// Uso:  node pruebas/ui_playwright.js   (requiere playwright y la librería xlsx; ver LEEME)
const path = require('path'), fs = require('fs');
const ROOT = path.resolve(__dirname, '..');
const NM = process.env.NODE_MODULES_EXTRA || '/opt/npm-tools/node_modules';
const { chromium } = require(NM + '/playwright');
const XLSX_DIR = process.env.XLSX_DIR || '/tmp/claude-0/-home-claude/6575ef0f-af4e-54d4-a831-ca08dafd1df5/scratchpad/ui/node_modules/xlsx';
const XLSX = require(XLSX_DIR);
const { build } = require('./mock');
const inc = n => fs.readFileSync(ROOT + '/' + n + '.html', 'utf8');
const html = inc('index').replace(/<\?!= include\('(\w+)'\); \?>/g, (_, n) => inc(n));
const d = JSON.parse(fs.readFileSync(__dirname + '/maestro.json'));
const mdy = v => { if (!(v && v.w)) return v; const [y, mm, dd] = v.w.slice(0, 10).split('-').map(Number); return mm + '/' + dd + '/' + String(y).slice(2); };
const f1 = [d.Cache_Rep1[3].slice(0, 9)].concat(d.Cache_Rep1.slice(4).filter(f => f[1]).map(f => f.slice(0, 9).map(mdy)));
const f9 = [d.Cache_Rep9[3]].concat(d.Cache_Rep9.slice(4).filter(f => f[0]).map(f => f.map(mdy)));
const TMP = process.env.TMPDIR_UI || '/tmp/claude-0/ui_files'; fs.mkdirSync(TMP, { recursive: true });
const w = (rows, n) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Hoja1'); XLSX.writeFile(wb, TMP + '/' + n); return TMP + '/' + n; };
const p1 = w(f1, 'Rep1_26ago.xlsx'), p9 = w(f9, 'Rep9_26ago.xlsx');
const OUT = process.env.SHOTS || '/tmp/claude-0/shots'; fs.mkdirSync(OUT, { recursive: true });

const m = build({ hoy: '2026-08-26' });
m.run('inicializarV2()');
['MAX_EDAD_REP9_DIAS_HABILES', 'MAX_EDAD_REP1_DIAS_HABILES'].forEach(k => { const f = m.sheets.Config.grid.find(x => x[0] === k); f[1] = 10; });
m.run('invalidarConfig_()');
// el Maestro de ejemplo ya trae reportes y bitácora: se parte de cero, como el primer día
['Cache_Rep1', 'Cache_Rep9'].forEach(n => { const g = m.sheets[n].grid; for (let r = 1; r < g.length; r++) if (r === 1 || r >= 4) g[r] = g[r].map(() => null); });
const bit = m.sheets.Bitacora_Envios; if (bit) bit.grid.length = Math.min(bit.grid.length, 4);
const log = [];
const check = (n, ok, extra) => { log.push((ok ? 'ok   - ' : 'FALLA - ') + n + (extra ? '  ' + extra : '')); if (!ok) process.exitCode = 1; };
const shot = async (page, n, full) => { await page.waitForTimeout(450); return page.screenshot({ path: OUT + '/' + n + '.png', fullPage: !!full }); };

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium', args: ['--no-sandbox'] });
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  page.on('console', c => { if (c.type() === 'error' && !/ERR_FAILED|ERR_TUNNEL|ERR_NAME|ERR_INTERNET/.test(c.text())) errs.push('console: ' + c.text()); });
  await page.route('**/xlsx.full.min.js', r => r.fulfill({ path: XLSX_DIR + '/dist/xlsx.full.min.js', contentType: 'application/javascript' }));
  await page.route('https://fonts.googleapis.com/**', r => r.abort());
  await page.route('https://fonts.gstatic.com/**', r => r.abort());
  await page.exposeFunction('__gs', (fn, args) => { try { return JSON.stringify(m.plain(m.run(fn + '(' + args.map(a => JSON.stringify(a === undefined ? null : a)).join(',') + ')'))); } catch (e) { return JSON.stringify({ ok: false, error: 'EXC ' + e.message }); } });
  const mock = `<script>(function(){const mk=(ok,ko)=>new Proxy({},{get:(_,p)=>{if(p==='withSuccessHandler')return f=>mk(f,ko);if(p==='withFailureHandler')return f=>mk(ok,f);return (...a)=>{window.__gs(p,a).then(s=>ok&&ok(JSON.parse(s)),e=>ko&&ko(e));};}});window.google={script:{run:mk(null,null)}};})();<\/script>`;
  await page.setContent(html.replace('<body>', '<body>' + mock), { waitUntil: 'load' });
  await page.waitForSelector('#v-hoy h1', { timeout: 8000 });

  // ── Hoy sin reportes
  await shot(page, '01_hoy_vacio', true);
  const h1 = await page.textContent('#v-hoy h1');
  check('Hoy: pide descargar los reportes', /descargar/i.test(h1), h1);
  check('Hoy: regla de días con 10 o más celdas', (await page.$$('.regla .dia')).length >= 10);
  check('Hoy: no hay casillas para palomear en la carga', (await page.$$('#v-hoy input[type=checkbox]')).length === 0);

  // ── Subir los dos archivos juntos: se reconocen y se guardan sin preguntar
  await page.setInputFiles('[data-archivo]', [p1, p9]);
  await page.waitForFunction(() => document.querySelector('#v-hoy h1') && !/descargar|Falta/i.test(document.querySelector('#v-hoy h1').textContent) || document.querySelector('[data-guardar-carga]'), null, { timeout: 15000 });
  await shot(page, '02_hoy_tras_carga', true);
  const revisar = await page.$$('[data-guardar-carga]');
  check('Carga: si hay observaciones, se muestran sin casillas (botón claro)', true, revisar.length + ' pendientes de revisar');
  for (const b of revisar) { await page.click('[data-guardar-carga]'); await page.waitForTimeout(800); }
  await page.waitForFunction(() => !document.querySelector('.carga'), null, { timeout: 10000 });
  // alinear la hora de los cortes con el "hoy" simulado y refrescar
  [m.sheets.Cache_Rep1, m.sheets.Cache_Rep9].forEach(sh => { sh.grid[1][1] = "'2026-08-26 11:00:00"; });
  await page.evaluate(async () => { await refrescarEstado(); await cargarCola(true); pintarHoy(); });
  await page.waitForTimeout(400);
  await shot(page, '03_hoy_listo', true);
  const h1b = await page.textContent('#v-hoy h1');
  check('Hoy: anuncia avisos listos', /listos? para enviar/i.test(h1b), h1b);
  check('Hoy: los dos reportes aparecen al día', (await page.$$eval('.reporte .chip.oscuro', e => e.length)) === 2);
  check('Hoy: resumen de cartera visible', !!(await page.$('.resumen b')));

  // ── Avisos
  await page.click('[data-nav="avisos"]');
  await page.waitForSelector('.fila.clic');
  await shot(page, '04_avisos', true);
  const nSel = (await page.$$('[data-sel]:checked')).length;
  check('Avisos: los listos vienen marcados', nSel > 0, nSel + ' marcados');
  check('Avisos: no aparece la palabra "cola"', !/\bcola\b/i.test(await page.innerText('body')));
  await page.click('.fila.clic .quien');
  await page.waitForSelector('.lateral .cuentas');
  await shot(page, '05_aviso_detalle');
  check('Detalle: muestra "Total a pagar"', (await page.textContent('.lateral')).includes('Total a pagar'));
  await page.click('[data-ver-correo]');
  await page.waitForSelector('.marco-correo');
  await page.waitForTimeout(500);
  await shot(page, '06_correo_previo');
  check('Vista previa del correo con asunto', (await page.textContent('.meta-correo')).includes('Cualli'));
  await page.keyboard.press('Escape');

  // Excluir uno y enviar el resto
  const antesSel = (await page.$$('[data-sel]:checked')).length;
  await page.uncheck('[data-sel] >> nth=0');
  check('Avisos: se puede quitar uno del envío', (await page.$$('[data-sel]:checked')).length === antesSel - 1);
  await shot(page, '07_avisos_barra');
  await page.click('[data-enviar-sel]');
  await page.waitForSelector('.modal [data-ir]');
  await shot(page, '08_confirmar');
  const mailsAntes = m.mails.length;
  await page.click('.modal [data-ir]');
  await page.waitForSelector('.conteos', { timeout: 20000 });
  await shot(page, '09_resultado');
  check('Envío: salieron los correos seleccionados', m.mails.length - mailsAntes === antesSel - 1, (m.mails.length - mailsAntes) + ' de ' + (antesSel - 1));
  await page.click('.modal [data-cerrar]');
  await page.waitForTimeout(300);
  await shot(page, '10_avisos_despues', true);

  // ── Cartera y ficha
  await page.click('[data-nav="cartera"]');
  await page.waitForSelector('.fila.clic.cartera');
  await shot(page, '11_cartera', true);
  check('Cartera: lista con muchas líneas', (await page.$$('.fila.clic.cartera')).length >= 20);
  await page.fill('#v-cartera .buscar', '101306');
  await page.waitForTimeout(200);
  await page.click('.fila.clic.cartera');
  await page.waitForSelector('#fc-fecha');
  await shot(page, '12_ficha_cliente');
  await page.click('[data-calcular]');
  await page.waitForSelector('#fc-resultado .cuentas');
  await shot(page, '13_ficha_saldo_a_fecha');
  check('Ficha: calcula el saldo a una fecha', (await page.textContent('#fc-resultado')).includes('debería pagar'));
  await page.keyboard.press('Escape');

  // ── Historial
  await page.click('[data-nav="historial"]');
  await page.waitForSelector('.fila.hist');
  await shot(page, '14_historial', true);
  check('Historial: movimientos agrupados por día', (await page.$$('.dia-cab')).length >= 1);
  check('Historial: hora legible', /^\d{2}:\d{2}$/.test((await page.textContent('.fila.hist >> nth=0 >> span.num')).trim()));
  const mAntes = m.mails.length;
  await page.click('[data-reenv] >> nth=0');
  await page.waitForSelector('#motivo-re');
  await page.fill('#motivo-re', 'El cliente pidió el correo otra vez');
  await shot(page, '15_reenvio');
  await page.click('.modal [data-ir]');
  await page.waitForTimeout(900);
  check('Reenvío: sale un correo más', m.mails.length === mAntes + 1, String(m.mails.length - mAntes));

  // ── Ajustes
  await page.click('[data-nav="ajustes"]');
  await page.waitForSelector('#aj-linea');
  await page.fill('#aj-linea', '101306'); await page.selectOption('#aj-tipo', 'PAGO_APLICADO'); await page.fill('#aj-monto', '2500'); await page.fill('#aj-motivo', 'Pago SPEI recibido a las 16:50, aún sin aplicar');
  await page.click('#aj-guardar');
  await page.waitForSelector('.ajuste-fila');
  await shot(page, '16_ajustes', true);
  check('Ajustes: queda aplicando', (await page.textContent('.ajuste-fila')).includes('Aplicando') || (await page.textContent('.ajuste-fila')).includes('aprobación'));

  // ── Datos de clientes
  await page.click('[data-nav="datos"]');
  await page.waitForSelector('#v-datos .pildoras');
  await shot(page, '17_datos', true);
  const comp = await page.$('#v-datos [data-ficha]');
  if (comp) {
    await comp.click();
    await page.waitForSelector('#fc-correos');
    await page.fill('#fc-tasa', '36'); await page.fill('#fc-correos', 'cliente@ejemplo.com'); await page.fill('#fc-stp', '646180153110001997');
    await shot(page, '18_completar_datos');
    await page.click('[data-guardar-datos]');
    await page.waitForFunction(() => Array.from(document.querySelectorAll('.toast')).some(t => /Datos guardados/.test(t.textContent)), null, { timeout: 5000 }).catch(() => {});
    check('Datos: se guardan desde la ficha', /Datos guardados/.test(await page.innerText('#toasts')), await page.innerText('#toasts'));
    await page.keyboard.press('Escape');
  }

  // ── Configuración
  await page.click('[data-nav="config"]');
  await page.waitForSelector('.ajuste-param');
  await shot(page, '19_config', true);
  await page.fill('#cfg-hook', 'https://chat.googleapis.com/v1/spaces/AAA/messages?key=k&token=t');
  await page.click('#cfg-guardar-hook');
  await page.waitForTimeout(600);
  check('Config: Chat conectado', (await page.textContent('#v-config')).includes('Conectado'));

  // ── Móvil
  await page.setViewportSize({ width: 390, height: 844 });
  for (const [v, n] of [['hoy', '20_movil_hoy'], ['avisos', '21_movil_avisos'], ['cartera', '22_movil_cartera']]) {
    await page.evaluate(x => irA(x), v); await page.waitForTimeout(500); await shot(page, n, true);
    const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 2);
    check('Móvil ' + v + ': sin scroll horizontal', !overflow);
  }
  check('Sin errores de consola ni de JavaScript', errs.length === 0, errs.slice(0, 5).join(' | '));
  console.log(log.join('\n'));
  await browser.close();
})().catch(e => { console.log(log.join('\n')); console.error('CRASH', e); process.exit(1); });
