// Mocks de Apps Script respaldados por el Excel real (maestro.json)
const fs = require('fs');
const { makeContext, loadGs, Utilities, wallToInstant } = require('./load');

function build(opts) {
  opts = opts || {};
  const data = JSON.parse(fs.readFileSync(__dirname + '/maestro.json', 'utf8'));
  let tz = opts.sheetTz || 'America/Phoenix';          // inferencia: hoja en UTC-7
  const sheets = {};
  const mails = [], fetches = [], triggers = [], props = {};

  const toCell = v => v;
  function Sheet(name, grid) { this.name = name; this.grid = grid || []; }
  const readCell = v => {
    if (v && typeof v === 'object' && v.w) { const m = v.w.match(/(\d+)-(\d+)-(\d+) (\d+):(\d+):(\d+)/); return wallToInstant(+m[1], +m[2], +m[3], +m[4], +m[5], +m[6], tz); }
    return v === undefined || v === null ? '' : v;
  };
  const writeCell = v => {
    if (Object.prototype.toString.call(v) === '[object Date]') return { w: Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss') };
    return v;
  };
  Sheet.prototype.getLastRow = function () {
    for (let i = this.grid.length - 1; i >= 0; i--) if ((this.grid[i] || []).some(c => c !== null && c !== undefined && c !== '')) return i + 1;
    return 0;
  };
  Sheet.prototype.getMaxColumns = function () { return 60; };
  Sheet.prototype.getRange = function (r, c, nr, nc) {
    const sh = this; nr = nr || 1; nc = nc || 1;
    if (typeof r === 'string') { // A1 notation simple: 'B2'
      const m = r.match(/^([A-Z]+)(\d+)$/); let col = 0; for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64; r = +m[2]; c = col; nr = 1; nc = 1;
    }
    const rng = {
      getValues() { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push(readCell((sh.grid[r - 1 + i] || [])[c - 1 + j])); out.push(row); } return out; },
      getValue() { return rng.getValues()[0][0]; },
      setValues(vals) { vals.forEach((row, i) => { while (sh.grid.length < r + i) sh.grid.push([]); const g = sh.grid[r - 1 + i] = sh.grid[r - 1 + i] || []; row.forEach((v, j) => { while (g.length < c - 1 + j) g.push(null); g[c - 1 + j] = writeCell(v); }); }); return rng; },
      setValue(v) { return rng.setValues([[v]]); },
      clearContent() { for (let i = 0; i < nr; i++) { const g = sh.grid[r - 1 + i]; if (g) for (let j = 0; j < nc; j++) g[c - 1 + j] = null; } return rng; },
      setFormula() { return rng; }
    };
    return new Proxy(rng, { get(t, p) { if (p in t) return t[p]; return () => new Proxy(rng, this); } });
  };
  Sheet.prototype.insertSheet = null;
  Sheet.prototype.setRowHeight = Sheet.prototype.setColumnWidth = Sheet.prototype.setFrozenRows = function () {};
  Object.keys(data).forEach(n => sheets[n] = new Sheet(n, data[n]));

  const ss = {
    getSheetByName: n => sheets[n] || null,
    insertSheet: n => (sheets[n] = new Sheet(n, [])),
    getSpreadsheetTimeZone: () => tz,
    setSpreadsheetTimeZone: z => { tz = z; }
  };
  const ctx = makeContext({
    SpreadsheetApp: { openById: () => ss },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; }, deleteProperty: k => { delete props[k]; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    MailApp: { sendEmail: (to, subj, plain, adv) => { if (opts.failTo && to.indexOf(opts.failTo) >= 0) throw new Error('Servicio no disponible'); mails.push({ to, subj, plain, adv }); }, getRemainingDailyQuota: () => opts.cuota === undefined ? 1500 : opts.cuota },
    Session: { getActiveUser: () => ({ getEmail: () => opts.usuario || 'robertodelacruz@cualli.mx' }) },
    UrlFetchApp: { fetch: (u, o) => { fetches.push({ u, o }); return { getResponseCode: () => 200 }; } },
    ScriptApp: { getProjectTriggers: () => triggers, deleteTrigger: t => { triggers.splice(triggers.indexOf(t), 1); },
      newTrigger: fn => { const t = { fn, getHandlerFunction: () => fn }; const b = { timeBased: () => b, everyDays: () => b, atHour: () => b, nearMinute: () => b, inTimezone: () => b, create: () => { triggers.push(t); return t; } }; return b; },
      getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/XXXX/exec' }) },
    Logger: { log: () => {} },
    HtmlService: {},
    __HOY__: opts.hoy || ''
  });
  loadGs(ctx, ['Config.gs', 'Calendario.gs', 'Tandas.gs', 'Datos.gs', 'Motor.gs', 'Email.gs', 'Sender.gs', 'Chat.gs', 'Setup.gs', 'Cartera.gs', 'Api.gs']);
  const run = s => require('vm').runInContext(s, ctx);
  return { ctx, run, sheets, mails, fetches, triggers, props, ss, setTz: z => { tz = z; }, plain: o => JSON.parse(JSON.stringify(o)) };
}
module.exports = { build };
