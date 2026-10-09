/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Config.gs
 *  Financiera Cualli SAPI de CV SOFOM ENR
 * ═══════════════════════════════════════════════════════════════════════════
 *  Constantes, lectura de la hoja Config (una sola lectura por ejecución),
 *  helpers de fecha tolerantes a zona horaria y control de acceso por rol.
 */

const SPREADSHEET_ID = '16SwfLDtLKsVsRPJd0ZdWKCH3J7Lt6GAdG1FACDLiwbY';
const TZ = 'America/Mexico_City';
const VERSION = '2.0.0';

const SHEETS = {
  CONFIG: 'Config',
  TASAS: 'Tasas',
  CORREOS: 'Correos',
  CACHE_REP1: 'Cache_Rep1',
  CACHE_REP9: 'Cache_Rep9',
  BITACORA: 'Bitacora_Envios',
  CORTES: 'Cortes',
  AJUSTES: 'Ajustes',
  INHABILES: 'Calendario_Inhabiles',
  USUARIOS: 'Usuarios',
  CAMBIOS: 'Cambios_Catalogo'
};

// Parámetros nuevos de la v2. inicializarV2() los agrega a Config SOLO si no existen.
const CONFIG_DEFAULTS = [
  ['TANDA_DIAS_HABILES', 1, 'Cada cuántos días hábiles se descarga un Rep1 nuevo (1 = todos los días hábiles).'],
  ['TANDA_ANCHOR_FECHA', '', 'Fecha aaaa-mm-dd de un día hábil en que empezó una tanda. Solo aplica si TANDA_DIAS_HABILES es mayor a 1.'],
  ['BASE_MORATORIOS', 'NOMINAL', 'Hasta qué fecha se proyectan los moratorios: NOMINAL (fecha del Rep1) o EFECTIVA (día hábil de pago).'],
  ['MAX_EDAD_REP9_DIAS_HABILES', 0, 'Antigüedad máxima del Rep9 para poder enviar, en días hábiles (0 = debe ser de hoy).'],
  ['MAX_EDAD_REP1_DIAS_HABILES', '', 'Antigüedad máxima del Rep1 en días hábiles. Vacío = TANDA_DIAS_HABILES - 1.'],
  ['ENVIO_AUTOMATICO', false, 'TRUE = un trigger envía solo las cuotas LISTAS a HORA_ENVIO_AUTOMATICO. FALSE = solo se envía desde la plataforma.'],
  ['HORA_RECORDATORIO', 8, 'Hora (0-23) del recordatorio por Chat para descargar reportes.'],
  ['HORA_ESCALACION', 9, 'Hora (0-23) en que se avisa por Chat si todavía no se cargaron los reportes.'],
  ['HORA_ENVIO_AUTOMATICO', 10, 'Hora (0-23) del envío automático (solo si ENVIO_AUTOMATICO = TRUE).'],
  ['REPLY_TO', '', 'Correo al que responde el cliente. Vacío = NOTIFICAR_A, y si tampoco existe, la cuenta que despliega.'],
  ['MODO_PRUEBA_DESTINO', '', 'En modo prueba todos los correos llegan aquí. Vacío = la cuenta de quien envía.'],
  ['FIRMA_NOMBRE', 'Karelia Monroy', 'Nombre que firma el correo.'],
  ['UMBRAL_AJUSTE_APROBACION', 50000, 'Ajustes de saldo por este monto o más requieren aprobación de otra persona (solo si hay APROBADORES).'],
  ['APROBADORES', '', 'Correos (separados por coma) que pueden aprobar ajustes grandes. Vacío = sin aprobación.'],
  ['CHAT_MENCION', '', 'A quién mencionar en el Chat: all, o el ID de usuario (users/123...). Vacío = sin mención.'],
  ['ASUNTO_RFC2047', true, 'TRUE conserva la codificación manual del asunto que ya funcionaba en la v1.']
];

// Valores por omisión de parámetros que ya existían en la v1.
const LEGACY_DEFAULTS = {
  FACTOR_TASA_MORATORIA: 2,
  DIAS_TIPO_T_MENOS_5: 5,
  DIAS_TIPO_T_MENOS_1: 1,
  SEPARACION_MINIMA_DIAS: 2,
  MODO_PRUEBA: false,
  REMITENTE_NOMBRE: 'Cobranza Cualli'
};

const BASE_DIAS_ANIO = 360;

// ─── Hoja de cálculo y configuración ───────────────────────────────────────

var _ss = null;
function spreadsheet_() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}

var _ssTz = null;
function ssTz_() {
  if (!_ssTz) _ssTz = spreadsheet_().getSpreadsheetTimeZone() || TZ;
  return _ssTz;
}

var _cfg = null;
function cfg_() {
  if (_cfg) return _cfg;
  const map = {};
  Object.keys(LEGACY_DEFAULTS).forEach(k => map[k] = LEGACY_DEFAULTS[k]);
  CONFIG_DEFAULTS.forEach(d => map[d[0]] = d[1]);
  const sh = spreadsheet_().getSheetByName(SHEETS.CONFIG);
  if (sh && sh.getLastRow() >= 3) {
    sh.getRange(3, 1, sh.getLastRow() - 2, 2).getValues().forEach(r => {
      const k = String(r[0] || '').trim();
      if (k) map[k] = r[1];
    });
  }
  _cfg = map;
  return map;
}
function invalidarConfig_() { _cfg = null; }

function cfgStr_(k, def) {
  const v = cfg_()[k];
  return (v === undefined || v === null || String(v).trim() === '') ? (def === undefined ? '' : def) : String(v).trim();
}
function cfgNum_(k, def) {
  const v = cfg_()[k];
  if (v === undefined || v === null || String(v).trim() === '') return def;
  const n = Number(v);
  return isNaN(n) ? def : n;
}
function cfgBool_(k, def) {
  const v = cfg_()[k];
  if (typeof v === 'boolean') return v;
  if (v === undefined || v === null || String(v).trim() === '') return def;
  const s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'VERDADERO' || s === '1' || s === 'SI' || s === 'SÍ';
}

// ─── Fechas: todo se maneja como texto aaaa-mm-dd ──────────────────────────
// Evita los corrimientos de un día entre la zona del script y la de la hoja.

function pad2_(n) { return String(n).padStart(2, '0'); }

function parseKey_(k) {
  const p = String(k).split('-');
  return new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
}
function keyOf_(d) {
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}
function addDays_(k, n) {
  const d = parseKey_(k);
  d.setUTCDate(d.getUTCDate() + n);
  return keyOf_(d);
}
function dow_(k) { return parseKey_(k).getUTCDay(); }              // 0 = domingo
function diffDays_(a, b) { return Math.round((parseKey_(b) - parseKey_(a)) / 86400000); } // b - a

/** Fecha de hoy en México. __HOY__ (pruebas) y FECHA_SIMULADA (solo modo prueba) la sustituyen. */
function hoyKey_() {
  if (typeof __HOY__ !== 'undefined' && __HOY__) return __HOY__;
  if (cfgBool_('MODO_PRUEBA', false)) {
    const s = PropertiesService.getScriptProperties().getProperty('FECHA_SIMULADA');
    if (s && /^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  }
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function fechaSimulada_() {
  if (typeof __HOY__ !== 'undefined' && __HOY__) return __HOY__;
  if (!cfgBool_('MODO_PRUEBA', false)) return '';
  return PropertiesService.getScriptProperties().getProperty('FECHA_SIMULADA') || '';
}

/**
 * Convierte lo que venga de una celda (Date, texto ISO, dd/mm/aaaa) a aaaa-mm-dd.
 * Un Date de la hoja se lee con la zona de la HOJA; si la hora es >= 18:00 se
 * interpreta como medianoche del día siguiente (así se corrige el desfase de
 * 1 hora histórico: fechas guardadas a las 23:00 del día anterior).
 */
function fechaKeyDeHoja_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return null;
    const s = Utilities.formatDate(v, ssTz_(), 'yyyy-MM-dd HH');
    let k = s.slice(0, 10);
    if (Number(s.slice(11, 13)) >= 18) k = addDays_(k, 1);
    return k;
  }
  const t = String(v).trim();
  let m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) {
    const y = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
    return y + '-' + pad2_(Number(m[2])) + '-' + pad2_(Number(m[1]));
  }
  return null;
}

/** Date al mediodía en la zona de la hoja: es lo que se escribe en las celdas de fecha. */
function fechaDeKey_(k) {
  return Utilities.parseDate(k + ' 12:00', ssTz_(), 'yyyy-MM-dd HH:mm');
}

function ahoraMX_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

// ─── Usuarios y permisos ───────────────────────────────────────────────────

const PERMISOS = {
  cargar:   ['COORDINADORA', 'GERENTE', 'ADMIN'],
  enviar:   ['COORDINADORA', 'GERENTE', 'ADMIN'],
  ajustar:  ['COORDINADORA', 'GERENTE', 'ADMIN'],
  aprobar:  ['GERENTE', 'ADMIN'],
  catalogo: ['COORDINADORA', 'GERENTE', 'ADMIN'],
  tasas:    ['GERENTE', 'ADMIN'],
  config:   ['ADMIN']
};

var _usuario = null;
function usuarioActual_() {
  if (_usuario) return _usuario;
  let email = '';
  try { email = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  let rol = 'ADMIN';
  let abierto = true;     // sin hoja Usuarios con filas activas, todos son ADMIN
  try {
    const sh = spreadsheet_().getSheetByName(SHEETS.USUARIOS);
    if (sh && sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
        .filter(r => String(r[0] || '').trim() && String(r[3] === '' ? 'SI' : r[3]).toUpperCase() !== 'NO');
      if (rows.length) {
        abierto = false;
        const mine = rows.find(r => String(r[0]).trim().toLowerCase() === email);
        rol = mine ? String(mine[2] || 'AUDITORIA').trim().toUpperCase() : 'SIN_ACCESO';
      }
    }
  } catch (e) {}
  _usuario = { email: email || 'desconocido', rol: rol, abierto: abierto };
  return _usuario;
}

function puede_(permiso) {
  const u = usuarioActual_();
  return (PERMISOS[permiso] || []).indexOf(u.rol) >= 0;
}

function requerir_(permiso) {
  if (!puede_(permiso)) {
    throw new Error('Tu rol (' + usuarioActual_().rol + ') no permite esta acción.');
  }
}
