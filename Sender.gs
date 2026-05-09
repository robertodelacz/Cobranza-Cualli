/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA — Envío y Bitácora (Sender.gs)
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─── API: PREVIEW (no envía) ───────────────────────────────────────────────

/**
 * Devuelve el HTML del correo para mostrar en modal de preview.
 * No envía nada ni registra en bitácora.
 */
function previsualizarCorreo(linea) {
  try {
    const aviso = obtenerAvisoPorLinea_(linea);
    if (!aviso) {
      return { ok: false, error: `Línea ${linea} no encontrada en los cálculos actuales.` };
    }

    const correo = construirCorreoAviso(aviso);

    // Validaciones previas
    const validacion = validarEnvio_(aviso);

    return {
      ok: true,
      asunto: correo.asunto,
      htmlBody: correo.htmlBody,
      destinatarios: correo.destinatarios,
      cc: leerCcFijo_(),
      from: Session.getActiveUser().getEmail(),
      validacion: validacion,
      aviso: aviso
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ─── API: ENVÍO INDIVIDUAL ─────────────────────────────────────────────────

/**
 * Envía 1 correo y registra en bitácora.
 */
function enviarAvisoIndividual(linea) {
  try {
    const aviso = obtenerAvisoPorLinea_(linea);
    if (!aviso) {
      return { ok: false, error: `Línea ${linea} no encontrada.` };
    }

    const validacion = validarEnvio_(aviso);
    if (!validacion.ok) {
      return { ok: false, error: validacion.error, validacion: validacion };
    }

    const resultado = ejecutarEnvio_(aviso);
    return resultado;

  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ─── API: ENVÍO MASIVO ─────────────────────────────────────────────────────

/**
 * Envía N correos secuencialmente. Si uno falla, los demás continúan.
 * @param {Array<string>} lineas - lista de líneas a enviar
 * @return {Object} reporte con resumen y detalle por línea
 */
function enviarAvisosMasivo(lineas) {
  const inicio = new Date();
  const resultados = [];
  let enviados = 0;
  let conError = 0;
  let omitidos = 0;

  for (const linea of lineas) {
    try {
      const aviso = obtenerAvisoPorLinea_(linea);
      if (!aviso) {
        resultados.push({ linea, ok: false, error: 'No encontrada en cálculos' });
        omitidos++;
        continue;
      }

      const validacion = validarEnvio_(aviso);
      if (!validacion.ok) {
        registrarBitacora_(aviso, 'OMITIDO', validacion.error);
        resultados.push({ linea, ok: false, error: validacion.error });
        omitidos++;
        continue;
      }

      const r = ejecutarEnvio_(aviso);
      resultados.push({ linea, ok: r.ok, error: r.error || null, cliente: aviso.nombre });
      if (r.ok) enviados++;
      else conError++;

      // Throttle suave: 200ms entre envíos para evitar spam-flag
      Utilities.sleep(200);

    } catch (err) {
      resultados.push({ linea, ok: false, error: err.message });
      conError++;
    }
  }

  const dur = Math.round((new Date().getTime() - inicio.getTime()) / 1000);

  return {
    ok: true,
    total: lineas.length,
    enviados: enviados,
    conError: conError,
    omitidos: omitidos,
    duracionSeg: dur,
    resultados: resultados
  };
}

// ─── EJECUCIÓN INTERNA DEL ENVÍO ───────────────────────────────────────────

/**
 * Ejecuta GmailApp.sendEmail y registra en bitácora.
 * Asume que ya pasó la validación.
 */
function ejecutarEnvio_(aviso) {
  const correo = construirCorreoAviso(aviso);
  const cc = leerCcFijo_();
  const destinatariosStr = correo.destinatarios.join(',');

  try {
    GmailApp.sendEmail(
      destinatariosStr,
      correo.asunto,
      correo.plainBody,
      {
        htmlBody: correo.htmlBody,
        cc: cc,
        name: 'Financiera Cualli',
        replyTo: Session.getActiveUser().getEmail()
      }
    );

    registrarBitacora_(aviso, 'ENVIADO', 'OK');

    return {
      ok: true,
      asunto: correo.asunto,
      destinatarios: correo.destinatarios,
      cc: cc
    };
  } catch (err) {
    registrarBitacora_(aviso, 'ERROR', err.message);
    return { ok: false, error: err.message };
  }
}

// ─── VALIDACIONES ──────────────────────────────────────────────────────────

/**
 * Valida que el aviso pueda enviarse: tiene STP, tiene correo, no se ha enviado hoy.
 */
function validarEnvio_(aviso) {
  // 1. Cuenta STP
  if (!aviso.cuentaSTP || String(aviso.cuentaSTP).trim() === '') {
    return { ok: false, error: 'Sin cuenta STP en el catálogo Correos.' };
  }

  // 2. Correo destinatario
  const destinatarios = parsearDestinatarios_(aviso.correo);
  if (destinatarios.length === 0) {
    return { ok: false, error: 'Sin correo destinatario válido.' };
  }

  // 3. No haber enviado el mismo aviso hoy (anti-doble envío)
  if (yaEnviadoHoy_(aviso.linea, aviso.tipoAviso)) {
    return {
      ok: false,
      error: `Aviso ${aviso.tipoAviso} para esta línea ya fue enviado hoy. Si quieres reenviar, hazlo desde la bitácora.`
    };
  }

  // 4. Tipo de aviso elegible
  if (!aviso.elegibleAviso) {
    return { ok: false, error: 'Tipo de aviso fuera de rango (T-5, T-1, T+0).' };
  }

  return { ok: true };
}

/**
 * Revisa la bitácora para ver si la línea + tipo ya se enviaron hoy.
 */
function yaEnviadoHoy_(linea, tipoAviso) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEETS.BITACORA);
  if (!sh) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 3) return false;

  const data = sh.getRange(3, 1, lastRow - 2, 8).getValues();
  // Cols: 0:Timestamp 1:FechaVenc 2:TipoAviso 3:Linea 4:Cliente 5:Correos 6:Total 7:Status

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  for (const row of data) {
    const ts = row[0];
    if (!ts || !(ts instanceof Date)) continue;
    const tsDay = new Date(ts.getFullYear(), ts.getMonth(), ts.getDate());
    if (tsDay.getTime() !== hoy.getTime()) continue;

    if (String(row[3]) === String(linea) &&
        String(row[2]) === String(tipoAviso) &&
        String(row[7]) === 'ENVIADO') {
      return true;
    }
  }
  return false;
}

// ─── BITÁCORA ──────────────────────────────────────────────────────────────

/**
 * Registra una fila en Bitacora_Envios.
 * @param {Object} aviso - el aviso (puede ser parcial si hay error temprano)
 * @param {string} status - 'ENVIADO' | 'ERROR' | 'OMITIDO'
 * @param {string} mensaje - detalle u OK
 */
function registrarBitacora_(aviso, status, mensaje) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEETS.BITACORA);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    const newRow = lastRow < 2 ? 3 : lastRow + 1;
    const fechaVenc = aviso.fechaVenc ? new Date(aviso.fechaVenc) : '';

    sh.getRange(newRow, 1, 1, 9).setValues([[
      new Date(),
      fechaVenc,
      aviso.tipoAviso || '',
      aviso.linea || '',
      aviso.nombre || aviso.cliente || '',
      aviso.correo || '',
      aviso.total || 0,
      status,
      mensaje || ''
    ]]);

    // Formato de la nueva fila
    sh.getRange(newRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sh.getRange(newRow, 2).setNumberFormat('dd/mm/yyyy');
    sh.getRange(newRow, 7).setNumberFormat('$#,##0.00');

  } catch (err) {
    // Silencioso: si la bitácora falla no debe tumbar el envío
    Logger.log('Error registrando bitácora: ' + err.message);
  }
}

// ─── HELPERS DE ACCESO A DATOS ─────────────────────────────────────────────

/**
 * Re-ejecuta calcularAvisos() y devuelve el aviso para una línea específica.
 * (Recalculamos en cada envío para asegurar valores frescos del cache actual.)
 */
function obtenerAvisoPorLinea_(linea) {
  const result = calcularAvisos();
  if (!result.ok) return null;
  return result.avisos.find(a => String(a.linea) === String(linea)) || null;
}

/**
 * Lee el valor de CC_FIJO de la hoja Config.
 */
function leerCcFijo_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEETS.CONFIG);
    if (!sh) return '';
    const data = sh.getRange(3, 1, sh.getLastRow() - 2, 2).getValues();
    for (const row of data) {
      if (String(row[0]).trim() === 'CC_FIJO') {
        return String(row[1] || '').trim();
      }
    }
    return '';
  } catch (err) {
    return '';
  }
}

// ─── API: BITÁCORA PARA EL FRONT ───────────────────────────────────────────

/**
 * Lee los últimos N envíos de la bitácora para mostrarlos en la webapp.
 */
function leerBitacora(limit) {
  try {
    const max = Math.max(1, Math.min(500, limit || 50));
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEETS.BITACORA);
    if (!sh) return { ok: true, registros: [] };

    const lastRow = sh.getLastRow();
    if (lastRow < 3) return { ok: true, registros: [] };

    const startRow = Math.max(3, lastRow - max + 1);
    const numRows = lastRow - startRow + 1;
    const data = sh.getRange(startRow, 1, numRows, 9).getValues();

    const registros = data.map(r => ({
      timestamp: r[0] instanceof Date ? r[0].toISOString() : '',
      fechaVenc: r[1] instanceof Date ? r[1].toISOString() : '',
      tipoAviso: r[2] || '',
      linea: r[3] || '',
      cliente: r[4] || '',
      correos: r[5] || '',
      total: Number(r[6]) || 0,
      status: r[7] || '',
      mensaje: r[8] || ''
    })).reverse();  // más reciente primero

    return { ok: true, registros: registros };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}
