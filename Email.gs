/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA — Generador de Correo (Email.gs)
 * ═══════════════════════════════════════════════════════════════════════════
 * 
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
  const fechaVenc = new Date(aviso.fechaPagoReal || aviso.fechaVenc); 
  
  const fechaCorta = formatearFechaCorta_(fechaVenc);
  const fechaLarga = formatearFechaLarga_(fechaVenc);
  const fechaSlash = formatearFechaSlash_(fechaVenc);
  const moneda = (aviso.moneda || 'MXN').toUpperCase();

  const asunto = `PRUEBA NO ANTENDER/ Cualli / Aviso de Pago / 🗓️ Vencimiento ${fechaCorta} / Línea ${aviso.linea}`;
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
          <td style="padding:28px 36px 20px 36px; background-color: #515151; background: linear-gradient(135deg, #515151 0%, #FFFFFF 100%); border-bottom:1px solid ${COLOR.GRAY_200};">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td valign="middle" width="180">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                    <tr>
                      <td valign="middle" style="background-color: #FFFFFF; border-radius: 8px; padding: 10px 14px; border: 1px solid #E5E5E5; box-shadow: 0px 4px 10px rgba(0,0,0,0.06);">
                        <img src="${LOGO_URL}" alt="Cualli" width="120" style="display:block; width:120px; max-width:120px; height:auto; border:0;">
                      </td>
                    </tr>
                  </table>
                </td>
<!-- Aviso de Pago y Fecha de Vencimiento  -->
                <td align="right" valign="middle">
                  <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="right">
                    <tr>
                      <td align="right" style="padding-bottom: 8px;">
                        
                        <!-- Badge amarillo (Abajo) -->
                        <div style="display: inline-block; background-color: ${COLOR.YELLOW}; color: ${COLOR.GRAY_INST}; padding: 5px 12px; border-radius: 6px; font-family: Arial, sans-serif; font-size: 11px; font-weight: bold; letter-spacing: 0.08em; text-transform: uppercase; box-shadow: 0px 3px 0px #DDA00C;">
                          <span style="font-size: 12px; margin-right: 4px; vertical-align: middle;">&#128276;</span> AVISO DE PAGO
                        </div>
                      </td>
                    </tr>
                  </table>
                </td>
                    </tr>
            </table>
          </td>
        </tr>
        <!-- ACENTO AMARILLO CON GRADIENTE -->
        <tr>
          <td style="background: linear-gradient(90deg, #FDB913 0%, #FFD66B 50%, #FDB913 100%); height: 4px; line-height: 0; font-size: 0;">&nbsp;</td>
        </tr>
        

<!-- SALUDO Y TEXTO -->
        <tr>
          <td style="padding:22px 36px 0 36px;">
            <p style="font-family: Arial, sans-serif; margin:0 0 14px 0; font-size:14px; color:${COLOR.GRAY_INST}; line-height:1.55;">
              Estimado Cliente: <strong style="color:${COLOR.GRAY_DARK};">${nombre}</strong>
            </p>
            <p style="font-family: Arial, sans-serif; margin:0 0 22px 0; font-size:14px; color:${COLOR.GRAY_INST}; line-height:1.55; text-align: justify;">
               Por medio del presente le recordamos que el que el próximo <strong style="color:${COLOR.GRAY_DARK};">${escaparHtml_(fechaLarga)}</strong> le corresponde realizar el pago referente a su línea de crédito No. <strong style="color:${COLOR.GRAY_DARK};">${linea}</strong>, por el importe que se detalla a continuación:
            </p>
          </td>
        </tr>
        <!-- CAJA DE MONTO ESTILIZADA  -->
        <tr>
          <td style="padding:10px 36px 30px 36px;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:${COLOR.GRAY_100}; border:1px solid ${COLOR.GRAY_200}; border-radius:8px;">
              <tr>
                <td align="center" style="padding:24px 20px;">
                  <div style="font-family: Arial, sans-serif; font-size:14px; color:${COLOR.GRAY_500}; letter-spacing:0.04em; font-weight:bold; text-transform:uppercase; margin-bottom:6px;">
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
            <div style="font-family: Arial, sans-serif; font-size:14px; color:#000000; font-weight:bold; letter-spacing:0.05em; text-transform:uppercase; margin-bottom:10px; padding-bottom:8px; border-bottom:2px solid ${COLOR.YELLOW};">
              Cuenta de depósito
            </div>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="font-family: Arial, sans-serif; font-size:14px; margin-bottom:20px; color:#000000;">
              <tr>
                <td style="font-weight:bold; padding:6px 0; width:30%;">Banco:</td>
                <td>${BANCO_FIJO}</td>
              </tr>
              <tr>
                <td style="font-weight:bold; padding:6px 0;">Beneficiario:</td>
                <td>${BENEFICIARIO_FIJO}</td>
              </tr>
              <tr>
                <td style="font-weight:bold; padding:6px 0; vertical-align:top;">CLABE:</td>
                <td style="font-family: Arial, sans-serif; font-weight:bold; letter-spacing:0.04em;">${stp}</td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- NOTA SOBRE HORARIO -->
        <tr>
          <td style="padding:0 36px 22px 36px;">
            <p style="font-family: Arial, sans-serif; margin:0; font-size:14px; color:${COLOR.GRAY_INST}; line-height:1.55; padding-top:14px; border-top:2px solid ${COLOR.YELLOW}; text-align: justify;">
             Es importante realizar su pago en tiempo y forma para mantener su cuenta al corriente y evitar la generación de intereses moratorios y/o comisiones. Considere que la hora límite de recepción de pagos a fin de mes es a las <strong style="color:#000000;">5:00 pm</strong>; los depositos realizados después de esta hora se aplicarán con fecha del día hábil siguiente.
            </p>
          </td>
        </tr>


        <!-- CIERRE -->
        <tr>
          <td style="padding:0 36px 15px 36px;">
            <p style="font-family: Arial, sans-serif; margin:0; font-size:13px; color:${COLOR.GRAY_INST}; line-height:1.6; text-align: justify;">
              Agradecemos su confirmación de depósito por este medio. Cualquier duda o aclaración estamos a sus órdenes.
            </p>
          </td>
        </tr>

        <!-- FIRMA DE LA ENCARGADA -->
        <tr>
          <td style="padding:0px 36px 30px 36px;">
            <!-- Despedida -->
            <p style="font-family: Arial, sans-serif; margin: 0 0 6px 0; font-size: 13px; color: ${COLOR.GRAY_INST};">
              Atentamente,
            </p>
            <!-- Nombre -->
            <p style="font-family: Arial, sans-serif; margin: 0; font-size: 14px; font-weight: bold; color: ${COLOR.GRAY_DARK};">
              Karelia Monroy
            </p>
          </td>
        </tr>
        <!-- INICIO DEL FOOTER UNIFICADO Y PROFESIONAL -->
        <tr>
          <td style="padding: 25px 36px 25px 36px; background-color: #F8F9FA; border-top: 1px solid #E9ECEF; border-radius: 0 0 8px 8px;">
            <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
              
              <!-- Razón Social, Slogan y Sitio Web -->
              <tr>
                <td align="left" valign="middle">
                  <p style="font-family: Arial, sans-serif; margin: 0 0 3px 0; font-size: 12px; font-weight: bold; color: ${COLOR.GRAY_DARK};">
                    Financiera Cualli, S.A.P.I. de C.V. SOFOM E.N.R.
                  </p>
                  <!-- Slogan integrado -->
                  <div style="font-family: Arial, sans-serif; font-size: 7.5pt; font-weight: bold; letter-spacing: 0.02em;">
                    <span style="color: #fbb818;">acelerando</span><span style="color: #525352;">oportunidades</span>
                  </div>
                </td>
                <td align="right" valign="middle">
                  <a href="https://cualli.mx" target="_blank" style="font-family: Arial, sans-serif; font-size: 12px; color: ${COLOR.GRAY_INST}; text-decoration: none; font-weight: bold;">
                    cualli.mx
                  </a>
                </td>
              </tr>

              <!-- Línea divisoria muy sutil -->
              <tr>
                <td colspan="2" style="padding-top: 15px; border-bottom: 1px solid #DEE2E6;"></td>
              </tr>

              <!-- Aviso legal / Recordatorio -->
              <tr>
                <td colspan="2" style="padding-top: 12px; text-align: justify;">
                  <p style="font-family: Arial, sans-serif; margin: 0; font-size: 10px; color: #868E96; line-height: 1.5;">
                   Este mensaje ha sido generado automáticamente con fines informativos. Si usted ya realizó el pago correspondiente, le pedimos hacer caso omiso de este recordatorio
                  </p>
                </td>
              </tr>

            </table>
          </td>
        </tr>
        <!-- FIN DEL FOOTER -->
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
