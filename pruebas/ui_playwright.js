const path = require('path'), fs = require('fs');
const SP = path.resolve(__dirname, '..');
const { chromium } = require('/opt/npm-tools/node_modules/playwright');
const XLSX = require('./node_modules/xlsx');
const { build } = require(SP + '/test/mock');
process.chdir(SP + '/test');
const V2 = SP + '/v2';
const inc = n => fs.readFileSync(V2 + '/' + n + '.html', 'utf8');
let html = inc('index').replace(/<\?!= include\('(\w+)'\); \?>/g, (_, n) => inc(n));
const d = JSON.parse(fs.readFileSync(SP + '/test/maestro.json'));
const mdy = v => { if (!(v && v.w)) return v; const [y, m, dd] = v.w.slice(0, 10).split('-').map(Number); return m + '/' + dd + '/' + String(y).slice(2); };
const f1 = [d.Cache_Rep1[1]].concat(d.Cache_Rep1.slice(2).filter(f => f[1]).map(f => f.slice(0, 9).map(mdy)));
const f9 = [d.Cache_Rep9[3]].concat(d.Cache_Rep9.slice(4).filter(f => f[0]).map(f => f.map(mdy)));
fs.mkdirSync('/tmp/claude-0/ui_files', { recursive: true });
const w = (rows, n) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Hoja1'); XLSX.writeFile(wb, '/tmp/claude-0/ui_files/' + n); return '/tmp/claude-0/ui_files/' + n; };
const p1 = w(f1, 'Rep1_26ago.xlsx'), p9 = w(f9, 'Rep9_26ago.xlsx');
const OUT = SP + '/shots'; fs.mkdirSync(OUT, { recursive: true });

const m = build({ hoy: '2026-08-26' });
m.run('inicializarV2()');
['MAX_EDAD_REP9_DIAS_HABILES', 'MAX_EDAD_REP1_DIAS_HABILES'].forEach(k => { const f = m.sheets.Config.grid.find(x => x[0] === k); f[1] = 10; });
m.run('invalidarConfig_()');
let log = [];
const check = (n, ok, extra) => { log.push((ok ? 'ok   - ' : 'FALLA - ') + n + (extra ? '  ' + extra : '')); if (!ok) process.exitCode = 1; };

(async () => {
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium', args: ['--no-sandbox'] });
  const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await ctx.newPage();
  const errs = [];
  page.on('pageerror', e => errs.push('pageerror: ' + e.message));
  page.on('console', c => { if (c.type() === 'error' && !/ERR_FAILED|ERR_TUNNEL/.test(c.text())) errs.push('console: ' + c.text()); });
  await page.route('**/xlsx.full.min.js', r => r.fulfill({ path: './../ui/node_modules/xlsx/dist/xlsx.full.min.js', contentType: 'application/javascript' }));
  await page.route('https://fonts.googleapis.com/**', r => r.abort());
  await page.route('https://fonts.gstatic.com/**', r => r.abort());
  await page.exposeFunction('__gs', (fn, args) => { try { return JSON.stringify(m.plain(m.run(fn + '(' + args.map(a => JSON.stringify(a === undefined ? null : a)).join(',') + ')'))); } catch (e) { return JSON.stringify({ ok: false, error: 'EXC ' + e.message }); } });
  const mock = `<script>(function(){const mk=(ok,ko)=>new Proxy({},{get:(_,p)=>{if(p==='withSuccessHandler')return f=>mk(f,ko);if(p==='withFailureHandler')return f=>mk(ok,f);return (...a)=>{window.__gs(p,a).then(s=>ok&&ok(JSON.parse(s)),e=>ko&&ko(e));};}});window.google={script:{run:mk(null,null)}};})();<\/script>`;
  await page.setContent(html.replace('<body>', '<body>' + mock), { waitUntil: 'load' });
  await page.waitForSelector('.cabecera h1', { timeout: 8000 });
  await page.screenshot({ path: OUT + '/01_hoy_vacio.png', fullPage: true });
  check('Hoy: titulo de tanda', (await page.textContent('#v-hoy h1')).includes('descargar el Rep1'), await page.textContent('#v-hoy h1'));
  check('Hoy: regla con 14+ celdas', (await page.$$('.regla .dia')).length >= 10);

  // Subir Rep1 y Rep9
  await page.setInputFiles('[data-archivo="rep1"]', p1);
  await page.waitForTimeout(2500); await page.waitForSelector('[data-confirma="rep1"]', {timeout:3000});
  await page.screenshot({ path: OUT + '/02_rep1_validado.png', fullPage: true });
  check('Rep1: el boton guardar inicia deshabilitado', await page.isDisabled('[data-guardar="rep1"]'));
  await page.check('[data-confirma="rep1"]');
  check('Rep1: guardar se habilita al confirmar', !(await page.isDisabled('[data-guardar="rep1"]')));
  await page.click('[data-guardar="rep1"]');
  await page.waitForSelector('.toast');
  await page.setInputFiles('[data-archivo="rep9"]', p9);
  await page.waitForSelector('[data-confirma="rep9"]');
  await page.check('[data-confirma="rep9"]');
  await page.click('[data-guardar="rep9"]');
  await page.waitForSelector('.cola-resumen', { timeout: 8000 });
  // alinear la hora de los cortes al "hoy" simulado
  [m.sheets.Cache_Rep1, m.sheets.Cache_Rep9].forEach(sh => { sh.grid[1][1] = "'2026-08-26 11:00:00"; });
  await page.evaluate(() => irA('hoy'));
  await page.waitForSelector('.cola-resumen');
  await page.screenshot({ path: OUT + '/03_hoy_cargado.png', fullPage: true });
  check('Hoy: resumen de cola presente', (await page.textContent('.cola-resumen')).includes('aviso'));

  // Cola
  await page.click('[data-nav="cola"]');
  await page.waitForSelector('#cuerpo-cola tr.fila');
  await page.screenshot({ path: OUT + '/04_cola.png', fullPage: true });
  const nFilas = (await page.$$('#cuerpo-cola tr.fila')).length;
  check('Cola: filas del filtro "Por enviar"', nFilas > 0, nFilas + ' filas');
  await page.click('#cuerpo-cola tr.fila .cliente');
  await page.waitForSelector('tr.detalle');
  await page.screenshot({ path: OUT + '/05_cola_detalle.png', fullPage: true });
  check('Cola: detalle muestra desglose', (await page.textContent('tr.detalle')).includes('Total a pagar'));
  await page.click('[data-filtro="todas"]');
  await page.screenshot({ path: OUT + '/06_cola_todas.png', fullPage: true });
  // vista previa
  await page.click('[data-filtro="enviar"]');
  await page.click('#cuerpo-cola tr.fila [data-ver]');
  await page.waitForSelector('.marco-correo');
  await page.waitForTimeout(500);
  await page.screenshot({ path: OUT + '/07_vista_previa.png' });
  check('Vista previa: asunto Cualli', (await page.textContent('.meta-correo')).includes('Cualli'));
  await page.click('.modal-pie [data-cerrar]');
  // seleccionar todas las listas y enviar
  await page.check('#chk-todas');
  await page.waitForSelector('.barra-envio');
  await page.screenshot({ path: OUT + '/08_seleccion.png' });
  await page.click('[data-enviar-sel]');
  await page.waitForSelector('.modal [data-ir]');
  await page.screenshot({ path: OUT + '/09_confirmar_envio.png' });
  const mailsAntes = m.mails.length;
  await page.click('.modal [data-ir]');
  await page.waitForSelector('.conteos', { timeout: 15000 });
  await page.screenshot({ path: OUT + '/10_resultado_envio.png' });
  check('Envio: salieron correos', m.mails.length > mailsAntes, (m.mails.length - mailsAntes) + ' correos');
  await page.click('.modal-pie [data-cerrar]');
  await page.screenshot({ path: OUT + '/11_cola_despues.png', fullPage: true });

  // Bitacora
  await page.click('[data-nav="bitacora"]');
  await page.waitForSelector('#v-bitacora tbody tr');
  await page.screenshot({ path: OUT + '/12_bitacora.png', fullPage: true });
  check('Bitacora: filas', (await page.$$('#v-bitacora tbody tr')).length > 1);
  check('Bitacora: hora legible', /\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}/.test(await page.textContent('#v-bitacora tbody tr:first-child td:first-child')), await page.textContent('#v-bitacora tbody tr:first-child td:first-child'));

  // Reenvio desde bitacora
  await page.click('[data-reenv]');
  await page.waitForSelector('#motivo-re');
  await page.fill('#motivo-re', 'El cliente pidio el correo otra vez');
  await page.screenshot({ path: OUT + '/13_reenvio.png' });
  const mAntes = m.mails.length;
  await page.click('.modal [data-ir]');
  await page.waitForTimeout(800);
  check('Reenvio: sale un correo mas', m.mails.length === mAntes + 1, (m.mails.length - mAntes) + '');

  // Ajustes
  await page.click('[data-nav="ajustes"]');
  await page.waitForSelector('#aj-linea');
  await page.fill('#aj-linea', '101306'); await page.selectOption('#aj-tipo', 'PAGO_APLICADO'); await page.fill('#aj-monto', '2500'); await page.fill('#aj-motivo', 'Pago SPEI recibido a las 16:50, sin aplicar');
  await page.click('#aj-guardar');
  await page.waitForSelector('#v-ajustes tbody .chip');
  await page.screenshot({ path: OUT + '/14_ajustes.png', fullPage: true });
  check('Ajustes: queda aplicando', (await page.textContent('#v-ajustes tbody')).includes('Aplicando'));

  // Catalogos
  await page.click('[data-nav="catalogos"]');
  await page.waitForSelector('#v-catalogos .panel');
  await page.screenshot({ path: OUT + '/15_catalogos.png', fullPage: true });
  await page.click('[data-ficha]');
  await page.waitForSelector('#f-correos');
  await page.fill('#f-tasa', '36'); await page.fill('#f-correos', 'cliente@ejemplo.com'); await page.fill('#f-stp', '646180153110001997');
  await page.screenshot({ path: OUT + '/16_ficha.png' });
  await page.click('.modal [data-ir]');
  await page.waitForSelector('.toast');
  check('Catalogos: ficha guardada', true);

  // Config
  await page.click('[data-nav="config"]');
  await page.waitForSelector('.params');
  await page.screenshot({ path: OUT + '/17_config.png', fullPage: true });
  await page.fill('#cfg-hook', 'https://chat.googleapis.com/v1/spaces/AAA/messages?key=k&token=t');
  await page.click('#cfg-guardar-hook');
  await page.waitForTimeout(500);
  check('Config: webhook conectado', (await page.textContent('#v-config')).includes('Conectado'));
  await page.click('#cfg-triggers'); await page.waitForTimeout(500);

  // Movil
  await page.setViewportSize({ width: 390, height: 844 });
  for (const [v, n] of [['hoy', '18_movil_hoy'], ['cola', '19_movil_cola']]) { await page.evaluate(x => irA(x), v); await page.waitForTimeout(300); await page.screenshot({ path: OUT + '/' + n + '.png', fullPage: true }); }
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 2);
  check('Movil: sin scroll horizontal de pagina', !overflow);

  check('Sin errores de consola/JS', errs.length === 0, errs.slice(0, 5).join(' | '));
  console.log(log.join('\n'));
  await browser.close();
})().catch(e => { console.log(log.join('\n')); console.error('CRASH', e); process.exit(1); });
