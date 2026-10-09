/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Calendario.gs
 *  Días hábiles. Sábado, domingo e inhábiles bancarios no cuentan.
 * ═══════════════════════════════════════════════════════════════════════════
 *  Los inhábiles se calculan por regla para cualquier año (así el 1-ene-2027
 *  ya no cae como día hábil) y se pueden corregir desde la hoja
 *  Calendario_Inhabiles: Estado ACTIVO agrega un día, QUITAR lo elimina.
 *
 *  Reglas verificadas contra los 11 días CNBV 2026 (ver pruebas). Para años
 *  posteriores son cálculo por regla: confírmalos cuando la CNBV publique su
 *  calendario oficial.
 */

function pascua_(y) {   // Meeus / Jones / Butcher
  const a = y % 19, b = Math.floor(y / 100), c = y % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return y + '-' + pad2_(mes) + '-' + pad2_(dia);
}

function nthMonday_(y, mes, n) {
  let k = y + '-' + pad2_(mes) + '-01', cuenta = 0;
  for (let i = 0; i < 40; i++) {
    if (dow_(k) === 1) { cuenta++; if (cuenta === n) return k; }
    k = addDays_(k, 1);
  }
  return null;
}

function inhabilesPorRegla_(y) {
  const p = pascua_(y);
  return [
    [y + '-01-01', 'Año Nuevo'],
    [nthMonday_(y, 2, 1), 'Día de la Constitución'],
    [nthMonday_(y, 3, 3), 'Natalicio de Benito Juárez'],
    [addDays_(p, -3), 'Jueves Santo'],
    [addDays_(p, -2), 'Viernes Santo'],
    [y + '-05-01', 'Día del Trabajo'],
    [y + '-09-16', 'Independencia de México'],
    [y + '-11-02', 'Día de Muertos'],
    [nthMonday_(y, 11, 3), 'Revolución Mexicana'],
    [y + '-12-12', 'Día del Empleado Bancario'],
    [y + '-12-25', 'Navidad']
  ];
}

var _calSet = null, _calQuitar = null, _calYears = {}, _calNombres = {};

function cargarCalendario_() {
  _calSet = new Set(); _calQuitar = new Set(); _calYears = {}; _calNombres = {};
  try {
    const sh = spreadsheet_().getSheetByName(SHEETS.INHABILES);
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
        const k = fechaKeyDeHoja_(r[0]);
        if (!k) return;
        const est = String(r[3] || 'ACTIVO').trim().toUpperCase();
        if (est === 'QUITAR') _calQuitar.add(k);
        else { _calSet.add(k); _calNombres[k] = String(r[1] || ''); }
      });
    }
  } catch (e) { /* sin hoja: solo reglas */ }
}
function invalidarCalendario_() { _calSet = null; }

function esInhabilFecha_(k) {
  if (!_calSet) cargarCalendario_();
  const y = Number(k.slice(0, 4));
  if (!_calYears[y]) {
    _calYears[y] = true;
    inhabilesPorRegla_(y).forEach(x => { _calSet.add(x[0]); if (!_calNombres[x[0]]) _calNombres[x[0]] = x[1]; });
  }
  if (_calQuitar.has(k)) return false;
  return _calSet.has(k);
}

function esHabil_(k) {
  const d = dow_(k);
  return d !== 0 && d !== 6 && !esInhabilFecha_(k);
}

/** Motivo por el que un día no cuenta ('' si es hábil). */
function motivoInhabil_(k) {
  const d = dow_(k);
  if (d === 6) return 'Sábado';
  if (d === 0) return 'Domingo';
  if (esInhabilFecha_(k)) return _calNombres[k] || 'Inhábil';
  return '';
}

/** El mismo día si es hábil; si no, el siguiente hábil. Es la fecha de pago efectiva. */
function siguienteHabil_(k) {
  let c = k;
  for (let i = 0; i < 400 && !esHabil_(c); i++) c = addDays_(c, 1);
  return c;
}
/** El mismo día si es hábil; si no, el hábil anterior. */
function anteriorHabil_(k) {
  let c = k;
  for (let i = 0; i < 400 && !esHabil_(c); i++) c = addDays_(c, -1);
  return c;
}

/** Avanza n días hábiles (cuenta solo días hábiles posteriores a k). */
function sumarHabiles_(k, n) {
  let c = k, cuenta = 0;
  for (let i = 0; i < 4000 && cuenta < n; i++) {
    c = addDays_(c, 1);
    if (esHabil_(c)) cuenta++;
  }
  return c;
}
/** Retrocede n días hábiles (cuenta solo días hábiles anteriores a k). */
function restarHabiles_(k, n) {
  let c = k, cuenta = 0;
  for (let i = 0; i < 4000 && cuenta < n; i++) {
    c = addDays_(c, -1);
    if (esHabil_(c)) cuenta++;
  }
  return c;
}

/** Días hábiles d con a < d <= b. Negativo si b es anterior a a. */
function habilesEntre_(a, b) {
  if (a === b) return 0;
  if (b < a) return -habilesEntre_(b, a);
  let c = a, n = 0;
  for (let i = 0; i < 4000 && c < b; i++) {
    c = addDays_(c, 1);
    if (esHabil_(c)) n++;
  }
  return n;
}

/** Lista de inhábiles de un año (para mostrar en Configuración). */
function inhabilesDelAnio_(y) {
  if (!_calSet) cargarCalendario_();
  esInhabilFecha_(y + '-01-01');
  const out = [];
  for (let k = y + '-01-01'; k <= y + '-12-31'; k = addDays_(k, 1)) {
    const d = dow_(k);
    if (d !== 0 && d !== 6 && esInhabilFecha_(k)) out.push({ fecha: k, descripcion: _calNombres[k] || 'Inhábil' });
  }
  return out;
}
