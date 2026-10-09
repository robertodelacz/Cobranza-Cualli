// Carga los .gs en un contexto vm, con mocks mínimos de Apps Script
const vm = require('vm'), fs = require('fs'), path = require('path');
const DIR = process.env.V2_DIR || path.join(__dirname, '..');
function wallToInstant(y, mo, d, h, mi, s, tz) {
  let guess = Date.UTC(y, mo - 1, d, h, mi, s);
  for (let i = 0; i < 3; i++) {
    const parts = new Intl.DateTimeFormat('en-CA', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(new Date(guess)).reduce((a, p) => (a[p.type] = p.value, a), {});
    const shown = Date.UTC(+parts.year, +parts.month - 1, +parts.day, +parts.hour, +parts.minute, +parts.second);
    guess += (Date.UTC(y, mo - 1, d, h, mi, s) - shown);
  }
  return new Date(guess);
}
const Utilities = {
  formatDate(date, tz, fmt) {
    const p = new Intl.DateTimeFormat('en-CA', { timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' }).formatToParts(date).reduce((a, x) => (a[x.type] = x.value, a), {});
    return fmt.replace(/'T'/g, 'T').replace('yyyy', p.year).replace('MM', p.month).replace('dd', p.day).replace('HH', p.hour).replace('mm', p.minute).replace('ss', p.second);
  },
  parseDate(str, tz, fmt) {
    const m = str.match(/(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?/);
    return wallToInstant(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0), tz);
  },
  sleep() {}, base64Encode: s => Buffer.from(s).toString('base64'),
  newBlob: s => ({ getBytes: () => Buffer.from(s) }),
  getUuid: () => require('crypto').randomUUID()
};
function makeContext(extra) {
  const ctx = Object.assign({ console, Utilities, Math, Date, JSON, Set, Map, Object, Array, String, Number, isNaN, parseInt, parseFloat, RegExp, Error }, extra || {});
  vm.createContext(ctx);
  return ctx;
}
function loadGs(ctx, files) {
  files.forEach(f => {
    const code = fs.readFileSync(path.join(DIR, f), 'utf8');
    // const/let de nivel superior no quedan en el global del vm: se exportan a var
    const patched = code.replace(/^(const|let)\s+/gm, 'var ');
    vm.runInContext(patched, ctx, { filename: f });
  });
}
module.exports = { makeContext, loadGs, Utilities, wallToInstant };
