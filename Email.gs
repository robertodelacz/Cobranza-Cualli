/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA — Generador de Correo (Email.gs)
 * ═══════════════════════════════════════════════════════════════════════════
 *  HTML profesional bancario con identidad Cualli sobria.
 *  Compatible con Gmail (table-based, sin CSS moderno).
 *
 *  Diseño:
 *    - Fuente única: Arial
 *    - Header: logo Cualli + "Aviso de Cobro"
 *    - Monto destacado en caja gris con texto blanco
 *    - Soporta MXN y USD
 *    - Ancho: 720px (sobrio, profesional)
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─── DATOS FIJOS ───────────────────────────────────────────────────────────
const BANCO_FIJO = 'Sistema de Transferencias y Pagos STP';
const BENEFICIARIO_FIJO = 'Financiera Cualli SAPI de CV SOFOM ENR';
const LOGO_URL = 'https://cualli.mx/wp-content/uploads/2022/07/cualli-bl@3x.png';
const EMAIL_WIDTH = 720;

// ─── COLORES INSTITUCIONALES ───────────────────────────────────────────────
const COLOR = {
  YELLOW:       '#FDB913',
  GRAY_INST:    '#515151',
  GRAY_DARK:    '#2E2E2E',
  GRAY_500:     '#6B6B6B',
  GRAY_300:     '#C8C8C8',
  GRAY_200:     '#E5E5E5',
  GRAY_100:     '#F2F2F2',
  GRAY_50:      '#FAFAFA',
  WHITE:        '#FFFFFF'
};

// ─── API PRINCIPAL ─────────────────────────────────────────────────────────

function construirCorreoAviso(aviso) {
  const fechaVenc = new Date(aviso.fechaVenc);
  const fechaCorta = formatearFechaCorta_(fechaVenc);
  const fechaLarga = formatearFechaLarga_(fechaVenc);
  const fechaSlash = formatearFechaSlash_(fechaVenc);
  const moneda = (aviso.moneda || 'MXN').toUpperCase();

  const asunto = `Cualli/ Aviso de Cobro / Vencimiento ${fechaCorta} / Línea ${aviso.linea}`;
  const htmlBody = construirHTML_(aviso, fechaSlash, fechaLarga, moneda);
  const plainBody = construirTextoPlano_(aviso, fechaLarga, moneda);
  const destinatarios = parsearDestinatarios_(aviso.correo);

  return { asunto, htmlBody, plainBody, destinatarios };
}

// ─── HTML BUILDER ──────────────────────────────────────────────────────────

function construirHTML_(aviso, fechaSlash, fechaLarga, moneda) {
  const totalFmt = formatearMoney_(aviso.total, moneda);
  const nombre = escaparHtml_(aviso.nombre || aviso.cliente || '');
  const linea = escaparHtml_(String(aviso.linea));
  const stp = escaparHtml_(String(aviso.cuentaSTP || '—'));

  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Aviso de Cobro</title>
</head>
<body style="margin:0; padding:0; background-color:${COLOR.GRAY_50}; font-family: Arial, Helvetica, sans-serif; color:${COLOR.GRAY_INST};">

<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:${COLOR.GRAY_50}; padding:32px 0;">
  <tr>
    <td align="center">

      <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="${EMAIL_WIDTH}" style="max-width:${EMAIL_WIDTH}px; width:100%; background-color:${COLOR.WHITE}; border:1px solid ${COLOR.GRAY_200}; border-radius:8px; border-collapse: separate; border-spacing: 0; overflow:hidden;">

        <!-- HEADER: Logo + 'Aviso de Cobro' -->
        <tr>
          <td style="padding:28px 36px 20px 36px; border-bottom:1px solid ${COLOR.GRAY_200};">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <!-- Logo Cualli compacto -->
                <td valign="middle" width="180">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td valign="middle" style="border-radius:6px; padding:8px 14px;">
                        <img src="${LOGO_URL}" alt="Cualli" width="100" style="display:block; width:140px; max-width:140px; height:auto; border:0;">
                      </td>
                    </tr>
                  </table>
                </td>
                <!-- Aviso de Cobro -->
                <td align="right" valign="middle">
                  <div style="font-family: Arial, sans-serif; font-size:12px; color:${COLOR.GRAY_500}; letter-spacing:0.06em; text-transform:uppercase; font-weight:bold;">
                    Aviso de Cobro
                  </div>
                  <div style="font-family: Arial, sans-serif; font-size:13px; color:${COLOR.GRAY_INST}; margin-top:4px;">
                    Línea ${linea}
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- ACENTO AMARILLO -->
        <tr>
          <td style="background-color:${COLOR.YELLOW}; height:3px; line-height:0; font-size:0;">&nbsp;</td>
        </tr>

        <!-- VENCIMIENTO -->
        <tr>
          <td style="padding:28px 36px 0 36px;">
            <div style="font-family: Arial, sans-serif; font-size:12px; color:${COLOR.GRAY_500}; letter-spacing:0.05em; text-transform:uppercase; margin-bottom:6px;">
              Fecha de vencimiento
            </div>
            <div style="font-family: Arial, sans-serif; font-size:18px; color:${COLOR.GRAY_DARK}; font-weight:bold; padding-bottom:18px; border-bottom:1px solid ${COLOR.GRAY_200};">
              ${escaparHtml_(fechaLarga)}
            </div>
          </td>
        </tr>

        <!-- SALUDO Y TEXTO -->
        <tr>
          <td style="padding:22px 36px 0 36px;">
            <p style="font-family: Arial, sans-serif; margin:0 0 14px 0; font-size:14px; color:${COLOR.GRAY_INST}; line-height:1.65;">
              Estimado Cliente: <strong style="color:${COLOR.GRAY_DARK};">${nombre}</strong>
            </p>
            <p style="font-family: Arial, sans-serif; margin:0 0 22px 0; font-size:14px; color:${COLOR.GRAY_INST}; line-height:1.65;">
               Por medio del presente le recordamos que el <strong style="color:${COLOR.GRAY_DARK};">${escaparHtml_(fechaLarga)}</strong> vence su pago programado correspondiente a la línea de crédito <strong style="color:${COLOR.GRAY_DARK};">${linea}</strong>.
            </p>
          </td>
        </tr>

        <!-- CAJA DE MONTO ESTILIZADA  -->
        <tr>
          <td style="padding:10px 36px 30px 36px;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:${COLOR.GRAY_100}; border:1px solid ${COLOR.GRAY_200}; border-radius:8px;">
              <tr>
                <td align="center" style="padding:24px 20px;">
                  <div style="font-family: Arial, sans-serif; font-size:12px; color:${COLOR.GRAY_500}; letter-spacing:0.04em; text-transform:uppercase; margin-bottom:6px;">
                    Cantidad a pagar
                  </div>
                  <div style="font-family: Arial, sans-serif; font-size:24px; color:${COLOR.GRAY_INST}; font-weight:bold; line-height:1.2;">
                    ${escaparHtml_(totalFmt)} <span style="font-size:14px; color:${COLOR.GRAY_500}; font-weight:normal;">${escaparHtml_(moneda)}</span>
                  </div>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- CUENTA DE DEPÓSITO -->
        <tr>
          <td style="padding:0 36px 6px 36px;">
            <div style="font-family: Arial, sans-serif; font-size:12px; color:${COLOR.GRAY_500}; letter-spacing:0.05em; text-transform:uppercase; margin-bottom:10px; padding-bottom:6px; border-bottom:1px solid ${COLOR.GRAY_200};">
              Cuenta de depósito
            </div>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="font-family: Arial, sans-serif; font-size:13px; margin-bottom:20px;">
              <tr>
                <td style="color:${COLOR.GRAY_500}; padding:6px 0; width:30%;">Banco:</td>
                <td style="color:${COLOR.GRAY_DARK};">${BANCO_FIJO}</td>
              </tr>
              <tr>
                <td style="color:${COLOR.GRAY_500}; padding:6px 0;">Beneficiario:</td>
                <td style="color:${COLOR.GRAY_DARK};">${BENEFICIARIO_FIJO}</td>
              </tr>
              <tr>
                <td style="color:${COLOR.GRAY_500}; padding:6px 0; vertical-align:top;">CLABE:</td>
                <td style="color:${COLOR.GRAY_DARK}; font-family: Arial, sans-serif; font-weight:bold; letter-spacing:0.04em;">${stp}</td>
              </tr>
            </table>
          </td>
        </tr>


        <!-- NOTA SOBRE HORARIO -->
        <tr>
          <td style="padding:0 36px 22px 36px;">
            <p style="font-family: Arial, sans-serif; margin:0; font-size:12px; color:${COLOR.GRAY_500}; line-height:1.55; padding-top:14px; border-top:1px solid ${COLOR.GRAY_200};">
              Es importante contar con el pago en tiempo y forma para evitar la generación de intereses moratorios. La hora límite para pagos a fin de mes es las <strong style="color:${COLOR.GRAY_INST};">5:00 pm</strong>; después se aplica con fecha del día hábil siguiente.
            </p>
          </td>
        </tr>

        <!-- CIERRE -->
        <tr>
          <td style="padding:0 36px 28px 36px;">
            <p style="font-family: Arial, sans-serif; margin:0; font-size:13px; color:${COLOR.GRAY_INST}; line-height:1.6;">
              Agradecemos su confirmación de depósito por este medio. Cualquier duda o aclaración estamos a sus órdenes.
            </p>
          </td>
        </tr>

        <!-- FOOTER INSTITUCIONAL (banda gris con datos) -->
        <tr>
          <td style="background-color:${COLOR.GRAY_INST}; padding:16px 36px;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="font-family: Arial, sans-serif; font-size:12px; color:${COLOR.YELLOW}; font-weight:bold; letter-spacing:0.03em;">
                  Financiera Cualli SAPI de CV SOFOM ENR
                </td>
                <td align="right" style="font-family: Arial, sans-serif; font-size:11px; color:${COLOR.GRAY_300};">
                  cualli.mx
                </td>
              </tr>
            </table>
          </td>
        </tr>

       <!-- AVISO LEGAL SOFOM ENR  -->
        <tr>
          <td style="background-color:${COLOR.GRAY_50}; padding:14px 36px; border-top:1px solid ${COLOR.GRAY_200}; border-bottom-left-radius: 7px; border-bottom-right-radius: 7px;">
            <p style="font-family: Arial, sans-serif; font-size:10px; color:${COLOR.GRAY_500}; line-height:1.5; margin:0; text-align:justify;">
              <strong style="color:${COLOR.GRAY_INST};">Aviso legal:</strong> Financiera Cualli SAPI de CV SOFOM ENR no requiere autorización de la Secretaría de Hacienda y Crédito Público para su constitución y operación, y está sujeta a la supervisión de la Comisión Nacional Bancaria y de Valores (CNBV) únicamente en materia de prevención de operaciones con recursos de procedencia ilícita y financiamiento al terrorismo.
            </p>
            <p style="font-family: Arial, sans-serif; font-size:10px; color:${COLOR.GRAY_500}; line-height:1.5; margin:8px 0 0 0;">
              Este es un recordatorio automático de un pago próximo a vencer. Si tiene dudas sobre este aviso o ya realizó su pago, le agradecemos confirmarlo respondiendo a este correo.
            </p>
          </td>
        </tr>
      </table>

    </td>
  </tr>
</table>

</body>
</html>`;
}

// ─── PLAIN TEXT FALLBACK ───────────────────────────────────────────────────

function construirTextoPlano_(aviso, fechaLarga, moneda) {
  const nombre = aviso.nombre || aviso.cliente || '';
  const totalFmt = formatearMoney_(aviso.total, moneda);

  return [
    `AVISO DE COBRO`,
    ``,
    `Estimado Cliente: ${nombre}`,
    ``,
    `Por medio del presente le recordamos que el ${fechaLarga} vence su pago programado correspondiente a la línea de crédito ${aviso.linea}.`,
    ``,
    `Cantidad a pagar: ${totalFmt} ${moneda}`,
    ``,
    `Cuenta de depósito:`,
    `  Banco: ${BANCO_FIJO}`,
    `  Beneficiario: ${BENEFICIARIO_FIJO}`,
    `  CLABE: ${aviso.cuentaSTP || '—'}`,
    ``,
    `Es importante contar con el pago en tiempo y forma para evitar la generación de intereses moratorios. La hora límite para pagos a fin de mes es las 5:00 pm; después se aplica con fecha del día hábil siguiente.`,
    ``,
    `Agradecemos su confirmación de depósito por este medio. Cualquier duda o aclaración estamos a sus órdenes.`,
    ``,
    `--`,
    `Financiera Cualli SAPI de CV SOFOM ENR`,
    `cualli.mx`,
    ``,
    `AVISO LEGAL: Financiera Cualli SAPI de CV SOFOM ENR no requiere autorización de la Secretaría de Hacienda y Crédito Público para su constitución y operación, y está sujeta a la supervisión de la Comisión Nacional Bancaria y de Valores (CNBV) únicamente en materia de prevención de operaciones con recursos de procedencia ilícita y financiamiento al terrorismo.`,
    ``,
    `Este es un recordatorio automático de un pago próximo a vencer. Si tiene dudas sobre este aviso o ya realizó su pago, le agradecemos confirmarlo respondiendo a este correo.`
  ].join('\n');
}

// ─── HELPERS ───────────────────────────────────────────────────────────────

function escaparHtml_(s) {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function formatearMoney_(n, moneda) {
  const num = Number(n) || 0;
  const formatted = num.toLocaleString('es-MX', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  if (moneda === 'USD') return 'US$' + formatted;
  return '$' + formatted;
}

function formatearFechaCorta_(d) {
  const meses = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  return `${pad2_(d.getDate())}-${meses[d.getMonth()]}-${d.getFullYear()}`;
}

function formatearFechaLarga_(d) {
  const dias  = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const diaTxt = dias[d.getDay()];
  const diaCap = diaTxt.charAt(0).toUpperCase() + diaTxt.slice(1);
  return `${diaCap} ${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`;
}

function formatearFechaSlash_(d) {
  return `${pad2_(d.getDate())}/${pad2_(d.getMonth() + 1)}/${d.getFullYear()}`;
}

function pad2_(n) { return String(n).padStart(2, '0'); }

function parsearDestinatarios_(raw) {
  if (!raw) return [];
  return String(raw)
    .split(/[,;]/)
    .map(s => s.trim())
    .filter(s => s.length > 0 && /@/.test(s));
}
