/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA v2 — Email.gs
 *  Plantillas del aviso de pago. Dos variantes de contenido:
 *    A  cuenta al corriente: monto de la cuota.
 *    B  con saldo vencido: se desglosa lo vencido y los moratorios.
 * ═══════════════════════════════════════════════════════════════════════════
 */

const BANCO_FIJO = 'Sistema de Transferencias y Pagos STP';
const BENEFICIARIO_FIJO = 'Financiera Cualli SAPI de CV SOFOM ENR';
const LOGO_URL = 'https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png';
const EMAIL_WIDTH = 720;

const COLOR = {
  YELLOW: '#FDB913', GRAY_INST: '#515151', GRAY_DARK: '#2E2E2E', GRAY_500: '#6B6B6B',
  GRAY_300: '#C8C8C8', GRAY_200: '#E5E5E5', GRAY_100: '#F2F2F2', GRAY_50: '#FAFAFA', WHITE: '#FFFFFF'
};

const AVISO_LEGAL = 'Financiera Cualli SAPI de CV SOFOM ENR no requiere autorización de la Secretaría de Hacienda y Crédito Público para su constitución y operación, y está sujeta a la supervisión de la Comisión Nacional Bancaria y de Valores (CNBV) únicamente en materia de prevención de operaciones con recursos de procedencia ilícita y financiamiento al terrorismo.';

// ─── API ───────────────────────────────────────────────────────────────────

/**
 * Arma el correo de una cuota. opts: { montoAnterior, esReenvio, motivoReenvio }.
 * El destinatario real se decide al enviar (en modo prueba se redirige).
 */
function construirCorreo_(item, opts) {
  opts = opts || {};
  const moneda = (item.moneda || 'MXN').toUpperCase();
  const fechaLarga = formatearFechaLarga_(item.fechaPago);
  const fechaCorta = formatearFechaCorta_(item.fechaPago);
  const esVispera = item.accion === 'VISPERA' || (item.previo && !item.accion && opts.esReenvio && item.estado !== 'PROGRAMADA');
  const prefijo = esVispera ? 'Recordatorio de pago' : 'Aviso de Pago';
  const asunto = 'Cualli / ' + prefijo + ' / 🗓️ Vencimiento ' + fechaCorta + ' / Línea ' + item.linea;
  const actualizado = opts.montoAnterior !== undefined && opts.montoAnterior !== null && Math.abs(opts.montoAnterior - item.total) > 0.01;
  return {
    asunto: asunto,
    html: construirHTML_(item, fechaLarga, moneda, esVispera, actualizado ? opts.montoAnterior : null),
    plain: construirTextoPlano_(item, fechaLarga, moneda, esVispera, actualizado ? opts.montoAnterior : null)
  };
}

function filasDesglose_(item, moneda) {
  const d = item.desglose, f = (n) => formatearMoney_(n, moneda);
  const filas = [];
  if (d.capital) filas.push(['Capital', f(d.capital)]);
  if (d.intereses) filas.push(['Intereses', f(d.intereses)]);
  if (d.otros) filas.push(['Otros cargos', f(d.otros)]);
  if (d.iva) filas.push(['IVA', f(d.iva)]);
  if (d.capVencido) filas.push(['Capital vencido', f(d.capVencido)]);
  if (d.intVencidos) filas.push(['Intereses vencidos (con IVA)', f(d.intVencidos)]);
  const mor = round2_(d.moratoriosAcum + d.moratoriosProy);
  if (mor) filas.push(['Intereses moratorios (con IVA)', f(mor)]);
  if (d.ajustes) filas.push([d.ajustes < 0 ? 'Pagos o abonos aplicados' : 'Otros movimientos', f(d.ajustes)]);
  return filas;
}

// ─── HTML ──────────────────────────────────────────────────────────────────

function construirHTML_(item, fechaLarga, moneda, esVispera, montoAnterior) {
  const totalFmt = formatearMoney_(item.total, moneda);
  const nombre = escaparHtml_(item.cliente || item.nombreRep1 || '');
  const linea = escaparHtml_(String(item.linea));
  const stp = escaparHtml_(String(item.cuentaSTP || '—'));
  const firma = escaparHtml_(cfgStr_('FIRMA_NOMBRE', 'Karelia Monroy'));
  const conVencido = item.plantilla === 'B';

  const intro = esVispera
    ? 'Le enviamos un segundo recordatorio: el próximo <strong style="color:' + COLOR.GRAY_DARK + ';">' + escaparHtml_(fechaLarga) + '</strong> le corresponde realizar el pago de su línea de crédito No. <strong style="color:' + COLOR.GRAY_DARK + ';">' + linea + '</strong>, por el importe que se detalla a continuación:'
    : 'Por medio del presente le recordamos que el próximo <strong style="color:' + COLOR.GRAY_DARK + ';">' + escaparHtml_(fechaLarga) + '</strong> le corresponde realizar el pago de su línea de crédito No. <strong style="color:' + COLOR.GRAY_DARK + ';">' + linea + '</strong>, por el importe que se detalla a continuación:';

  const notaVencido = conVencido
    ? '<p style="font-family: Arial, sans-serif; margin:0 0 14px 0; font-size:13px; color:' + COLOR.GRAY_INST + '; line-height:1.55; text-align: justify;">Su línea presenta saldos vencidos; el importe incluye lo vencido y los intereses moratorios calculados a la fecha de pago.</p>'
    : '';
  const notaActualizado = montoAnterior !== null
    ? '<p style="font-family: Arial, sans-serif; margin:0 0 14px 0; font-size:13px; color:' + COLOR.GRAY_INST + '; line-height:1.55; text-align: justify;">El importe se actualizó respecto del aviso anterior (' + escaparHtml_(formatearMoney_(montoAnterior, moneda)) + ') por movimientos aplicados en su cuenta.</p>'
    : '';

  const desglose = filasDesglose_(item, moneda).map(function (f) {
    return '<tr><td style="padding:5px 0; color:' + COLOR.GRAY_500 + ';">' + escaparHtml_(f[0]) + '</td><td align="right" style="padding:5px 0; color:' + COLOR.GRAY_DARK + ';">' + escaparHtml_(f[1]) + '</td></tr>';
  }).join('');
  const bloqueDesglose = desglose
    ? '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-top:14px; border-top:1px solid ' + COLOR.GRAY_200 + '; font-family: Arial, sans-serif; font-size:13px;">' + desglose + '</table>'
    : '';

  return '<!DOCTYPE html>\n<html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Aviso de pago</title></head>\n' +
'<body style="margin:0; padding:0; background-color:' + COLOR.GRAY_50 + '; font-family: Arial, Helvetica, sans-serif; color:' + COLOR.GRAY_INST + ';">\n' +
'<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:' + COLOR.GRAY_50 + '; padding:32px 0;"><tr><td align="center">\n' +
'<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="' + EMAIL_WIDTH + '" style="max-width:' + EMAIL_WIDTH + 'px; width:100%; background-color:' + COLOR.WHITE + '; border:1px solid ' + COLOR.GRAY_200 + '; border-radius:8px; border-collapse: separate; border-spacing: 0; overflow:hidden;">\n' +
// Encabezado
'<tr><td style="padding:28px 36px 20px 36px; background-color:#515151; background: linear-gradient(135deg, #515151 0%, #FFFFFF 100%); border-bottom:1px solid ' + COLOR.GRAY_200 + ';">\n' +
'<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>\n' +
'<td valign="middle" width="180"><table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr><td valign="middle" style="background-color:#FFFFFF; border-radius:8px; padding:10px 14px; border:1px solid #E5E5E5;"><img src="' + LOGO_URL + '" alt="Cualli" width="120" style="display:block; width:120px; max-width:120px; height:auto; border:0;"></td></tr></table></td>\n' +
'<td align="right" valign="middle"><div style="display:inline-block; background-color:' + COLOR.YELLOW + '; color:' + COLOR.GRAY_INST + '; padding:5px 12px; border-radius:6px; font-family: Arial, sans-serif; font-size:11px; font-weight:bold; letter-spacing:0.08em; text-transform:uppercase; box-shadow:0px 3px 0px #DDA00C;">' + (esVispera ? 'RECORDATORIO DE PAGO' : 'AVISO DE PAGO') + '</div></td>\n' +
'</tr></table></td></tr>\n' +
'<tr><td style="background: linear-gradient(90deg, #FDB913 0%, #FFD66B 50%, #FDB913 100%); height:4px; line-height:0; font-size:0;">&nbsp;</td></tr>\n' +
// Saludo
'<tr><td style="padding:22px 36px 0 36px;">\n' +
'<p style="font-family: Arial, sans-serif; margin:0 0 14px 0; font-size:14px; color:' + COLOR.GRAY_INST + '; line-height:1.55;">Estimado Cliente: <strong style="color:' + COLOR.GRAY_DARK + ';">' + nombre + '</strong></p>\n' +
'<p style="font-family: Arial, sans-serif; margin:0 0 18px 0; font-size:14px; color:' + COLOR.GRAY_INST + '; line-height:1.55; text-align: justify;">' + intro + '</p>\n' +
notaVencido + notaActualizado +
'</td></tr>\n' +
// Monto
'<tr><td style="padding:6px 36px 28px 36px;">\n' +
'<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:' + COLOR.GRAY_100 + '; border:1px solid ' + COLOR.GRAY_200 + '; border-radius:8px;"><tr><td style="padding:22px 24px;">\n' +
'<div align="center" style="font-family: Arial, sans-serif; font-size:14px; color:' + COLOR.GRAY_500 + '; letter-spacing:0.04em; font-weight:bold; text-transform:uppercase; margin-bottom:6px;">Cantidad a pagar</div>\n' +
'<div align="center" style="font-family: Arial, sans-serif; font-size:26px; color:' + COLOR.GRAY_INST + '; font-weight:bold; line-height:1.2;">' + escaparHtml_(totalFmt) + ' <span style="font-size:14px; color:' + COLOR.GRAY_500 + '; font-weight:normal;">' + escaparHtml_(moneda) + '</span></div>\n' +
bloqueDesglose +
'</td></tr></table></td></tr>\n' +
// Cuenta
'<tr><td style="padding:0 36px 6px 36px;">\n' +
'<div style="font-family: Arial, sans-serif; font-size:14px; color:#000000; font-weight:bold; letter-spacing:0.05em; text-transform:uppercase; margin-bottom:10px; padding-bottom:8px; border-bottom:2px solid ' + COLOR.YELLOW + ';">Cuenta de depósito</div>\n' +
'<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="font-family: Arial, sans-serif; font-size:14px; margin-bottom:20px; color:#000000;">\n' +
'<tr><td style="font-weight:bold; padding:6px 0; width:30%;">Banco:</td><td>' + BANCO_FIJO + '</td></tr>\n' +
'<tr><td style="font-weight:bold; padding:6px 0;">Beneficiario:</td><td>' + BENEFICIARIO_FIJO + '</td></tr>\n' +
'<tr><td style="font-weight:bold; padding:6px 0; vertical-align:top;">CLABE:</td><td style="font-weight:bold; letter-spacing:0.04em;">' + stp + '</td></tr>\n' +
'</table></td></tr>\n' +
// Horario
'<tr><td style="padding:0 36px 22px 36px;"><p style="font-family: Arial, sans-serif; margin:0; font-size:14px; color:' + COLOR.GRAY_INST + '; line-height:1.55; padding-top:14px; border-top:2px solid ' + COLOR.YELLOW + '; text-align: justify;">Es importante realizar su pago en tiempo y forma para mantener su cuenta al corriente y evitar la generación de intereses moratorios y/o comisiones. Considere que la hora límite de recepción de pagos a fin de mes es a las <strong style="color:#000000;">5:00 pm</strong>; los depósitos realizados después de esta hora se aplicarán con fecha del día hábil siguiente.</p></td></tr>\n' +
// Cierre y firma
'<tr><td style="padding:0 36px 15px 36px;"><p style="font-family: Arial, sans-serif; margin:0; font-size:13px; color:' + COLOR.GRAY_INST + '; line-height:1.6; text-align: justify;">Agradecemos su confirmación de depósito por este medio. Cualquier duda o aclaración estamos a sus órdenes.</p></td></tr>\n' +
'<tr><td style="padding:0 36px 30px 36px;"><p style="font-family: Arial, sans-serif; margin:0 0 6px 0; font-size:13px; color:' + COLOR.GRAY_INST + ';">Atentamente,</p><p style="font-family: Arial, sans-serif; margin:0; font-size:14px; font-weight:bold; color:' + COLOR.GRAY_DARK + ';">' + firma + '</p></td></tr>\n' +
// Pie
'<tr><td style="padding:25px 36px; background-color:#F8F9FA; border-top:1px solid #E9ECEF;">\n' +
'<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>\n' +
'<td align="left" valign="middle"><p style="font-family: Arial, sans-serif; margin:0 0 3px 0; font-size:12px; font-weight:bold; color:' + COLOR.GRAY_DARK + ';">Financiera Cualli, S.A.P.I. de C.V. SOFOM E.N.R.</p><div style="font-family: Arial, sans-serif; font-size:10px; font-weight:bold; letter-spacing:0.02em;"><span style="color:#fbb818;">acelerando</span><span style="color:#525352;">oportunidades</span></div></td>\n' +
'<td align="right" valign="middle"><a href="https://cualli.mx" target="_blank" style="font-family: Arial, sans-serif; font-size:12px; color:' + COLOR.GRAY_INST + '; text-decoration:none; font-weight:bold;">cualli.mx</a></td>\n' +
'</tr><tr><td colspan="2" style="padding-top:15px; border-bottom:1px solid #DEE2E6;"></td></tr>\n' +
'<tr><td colspan="2" style="padding-top:12px; text-align: justify;"><p style="font-family: Arial, sans-serif; margin:0 0 8px 0; font-size:10px; color:#868E96; line-height:1.5;">Este mensaje se generó automáticamente con fines informativos. Si usted ya realizó el pago correspondiente, le pedimos hacer caso omiso de este recordatorio.</p><p style="font-family: Arial, sans-serif; margin:0; font-size:10px; color:#868E96; line-height:1.5;">' + escaparHtml_(AVISO_LEGAL) + '</p></td></tr>\n' +
'</table></td></tr>\n' +
'</table>\n</td></tr></table>\n</body></html>';
}

// ─── TEXTO PLANO ───────────────────────────────────────────────────────────

function construirTextoPlano_(item, fechaLarga, moneda, esVispera, montoAnterior) {
  const nombre = item.cliente || item.nombreRep1 || '';
  const lineas = [
    esVispera ? 'RECORDATORIO DE PAGO' : 'AVISO DE PAGO', '',
    'Estimado Cliente: ' + nombre, '',
    (esVispera ? 'Le enviamos un segundo recordatorio: el próximo ' : 'Por medio del presente le recordamos que el próximo ') + fechaLarga +
      ' le corresponde realizar el pago de su línea de crédito No. ' + item.linea + ', por el importe que se detalla a continuación.', ''
  ];
  if (item.plantilla === 'B') lineas.push('Su línea presenta saldos vencidos; el importe incluye lo vencido y los intereses moratorios calculados a la fecha de pago.', '');
  if (montoAnterior !== null) lineas.push('El importe se actualizó respecto del aviso anterior (' + formatearMoney_(montoAnterior, moneda) + ') por movimientos aplicados en su cuenta.', '');
  lineas.push('Cantidad a pagar: ' + formatearMoney_(item.total, moneda) + ' ' + moneda);
  filasDesglose_(item, moneda).forEach(f => lineas.push('  ' + f[0] + ': ' + f[1]));
  lineas.push('', 'Cuenta de depósito:', '  Banco: ' + BANCO_FIJO, '  Beneficiario: ' + BENEFICIARIO_FIJO, '  CLABE: ' + (item.cuentaSTP || '—'), '',
    'Es importante realizar su pago en tiempo y forma para evitar la generación de intereses moratorios y/o comisiones. La hora límite de recepción de pagos a fin de mes es las 5:00 pm; los depósitos posteriores se aplican con fecha del día hábil siguiente.', '',
    'Agradecemos su confirmación de depósito por este medio. Cualquier duda o aclaración estamos a sus órdenes.', '',
    'Atentamente,', cfgStr_('FIRMA_NOMBRE', 'Karelia Monroy'), '--', 'Financiera Cualli SAPI de CV SOFOM ENR', 'cualli.mx', '',
    'Este mensaje se generó automáticamente con fines informativos. Si usted ya realizó el pago correspondiente, le pedimos hacer caso omiso de este recordatorio.', '',
    'AVISO LEGAL: ' + AVISO_LEGAL);
  return lineas.join('\n');
}

// ─── HELPERS ───────────────────────────────────────────────────────────────

function escaparHtml_(s) {
  if (s === null || s === undefined) return '';
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
}

/** Formato $1,234.56 sin depender del idioma del servidor. */
function formatearMoney_(n, moneda) {
  const v = Number(n) || 0;
  const neg = v < 0;
  const partes = Math.abs(v).toFixed(2).split('.');
  const miles = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (neg ? '-' : '') + (moneda === 'USD' ? 'US$' : '$') + miles + '.' + partes[1];
}

const MESES_CORTO = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const MESES_LARGO = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS_LARGO = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

function formatearFechaCorta_(k) { return k.slice(8, 10) + '-' + MESES_CORTO[Number(k.slice(5, 7)) - 1] + '-' + k.slice(0, 4); }
function formatearFechaLarga_(k) {
  const dia = DIAS_LARGO[dow_(k)];
  return dia.charAt(0).toUpperCase() + dia.slice(1) + ' ' + Number(k.slice(8, 10)) + ' de ' + MESES_LARGO[Number(k.slice(5, 7)) - 1] + ' de ' + k.slice(0, 4);
}
