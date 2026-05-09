/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA — Backend (Code.gs)
 *  Financiera Cualli SAPI de CV SOFOM ENR · Contraloría
 * ═══════════════════════════════════════════════════════════════════════════
 *  Módulo de carga: Rep1 (vencimientos) y Rep9 (saldos) → Spreadsheet maestro
 *
 *  ESTRUCTURA REP1 (ACTUALIZADA):
 *    Se cargan SOLO las columnas B-K del archivo original (10 cols), omitiendo:
 *    - Col A: "Prod." (no se usa)
 *    - Col D: "Linea de Crédito" duplicada (queda solo la C)
 *    - Cols L-O: "Plan/Real", "Sucursal", "Promotor", "Cobrador" (no relevantes)
 *    - Cols al final amarillas (totales informativos)
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─── CONFIGURACIÓN GLOBAL ──────────────────────────────────────────────────
const SPREADSHEET_ID = '16SwfLDtLKsVsRPJd0ZdWKCH3J7Lt6GAdG1FACDLiwbY';

const SHEETS = {
  CONFIG: 'Config',
  TASAS: 'Tasas',
  CORREOS: 'Correos',
  CACHE_REP1: 'Cache_Rep1',
  CACHE_REP9: 'Cache_Rep9',
  BITACORA: 'Bitacora_Envios'
};

// ─── HEADERS ESPERADOS ─────────────────────────────────────────────────────

/**
 * Headers esperados en el archivo Rep1 que el usuario sube.
 * El parser detecta el inicio del header real (puede no estar en row 1) y
 * descarta la columna "Linea de Crédito" duplicada (suele ser col D).
 */
const REP1_RAW_HEADERS_EXPECTED = [
  'Fecha Vencimiento',
  'Linea de Crédito',
  'Nombre de Cliente',
  'Capital',
  'Interés',
  'Otros',
  'IVA',
  'Importe',
  'Moneda'
];

/**
 * Headers que se guardan en Cache_Rep1 (10 columnas finales).
 */
const REP1_HEADERS = [
  'Fecha Vencimiento',  // A
  'Línea de Crédito',   // B
  'Nombre de Cliente',  // C
  'Capital',            // D
  'Interés',            // E
  'Otros',              // F
  'IVA',                // G
  'Importe',            // H
  'Moneda'              // I
];

const REP9_HEADERS = [
  'Línea de crédito', 'Nombre Cliente', 'Num. Cliente_ID', 'Moneda',
  'Fecha Alta', 'Fijo hasta', '*Saldo disp. de la línea', '*Monto línea',
  'Monto del crédito', 'Cap. Vigente', 'Cap. Vencido', '*Saldo Cap.',
  'IVA Cliente', 'Int. Vigente', 'IVA Int. Vigente', 'Int. Vencido',
  'IVA Int. Vencido', 'Moratorios', 'IVA Moratorios', 'Mor. Contabilizados',
  'IVA Mor. Contabilizados', '*Saldo Int.', 'IVA Comisiones',
  '*Saldo Comisiones', '*IVA Saldo Comisiones', '*Saldo Vigente',
  '*Saldo Vencido', '*SALDO TOTAL', 'Pagos No Aplicados', 'Promotor'
];

// ─── ENTRY POINTS ──────────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cobranza Preventiva · Cualli')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── API PARA EL FRONT ─────────────────────────────────────────────────────

function getInitialState() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userEmail = Session.getActiveUser().getEmail() || 'desconocido';

    return {
      ok: true,
      userEmail: userEmail,
      rep1Status: getCacheStatus_(ss, SHEETS.CACHE_REP1),
      rep9Status: getCacheStatus_(ss, SHEETS.CACHE_REP9),
      now: new Date().toISOString()
    };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Valida y normaliza el Rep1 que llegó del cliente.
 * El cliente envía las filas RAW del archivo (todas las columnas tal cual).
 * Aquí detectamos el header, eliminamos la columna duplicada D (Linea repetida)
 * y truncamos a las 10 columnas relevantes.
 */
function validateRep1(rows) {
  const normalized = normalizarRep1Raw_(rows);
  if (!normalized.ok) return normalized;
  return {
    ok: true,
    label: 'Rep1',
    rowCount: normalized.dataRows.length,
    sample: normalized.dataRows.slice(0, 3),
    normalizedRows: normalized.allRows  // header + datos ya limpios
  };
}

function validateRep9(rows) {
  return validateReport_(rows, REP9_HEADERS, 'Rep9');
}

function saveRep1ToCache(rows) {
  const normalized = normalizarRep1Raw_(rows);
  if (!normalized.ok) return normalized;
  return saveCache_(normalized.allRows, SHEETS.CACHE_REP1, REP1_HEADERS, 'Rep1');
}

function saveRep9ToCache(rows) {
  return saveCache_(rows, SHEETS.CACHE_REP9, REP9_HEADERS, 'Rep9');
}

// ─── NORMALIZACIÓN REP1 ────────────────────────────────────────────────────

/**
 * Toma el Rep1 RAW del cliente y devuelve { allRows: [headers, ...datos] }
 * con las 10 columnas finales.
 *
 * Lógica:
 *   1. Encuentra la fila de headers buscando "Fecha Vencimiento"
 *   2. Identifica las columnas relevantes (Fecha Venc → Moneda)
 *   3. Detecta y descarta la columna "Línea de Crédito" duplicada
 *   4. Devuelve solo esas columnas, en orden estándar
 */
function normalizarRep1Raw_(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return { ok: false, error: 'Rep1: el archivo no contiene datos suficientes.' };
  }

  // 1. Localizar fila de headers
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const fila = rows[i].map(c => String(c || '').toLowerCase().trim());
    if (fila.some(c => c.includes('fecha vencimiento') || c === 'fecha venc.')) {
      headerRowIdx = i;
      break;
    }
  }

  if (headerRowIdx === -1) {
    return {
      ok: false,
      error: 'Rep1: no se encontró el header "Fecha Vencimiento" en las primeras 10 filas.'
    };
  }

  const rawHeaders = rows[headerRowIdx].map(c => String(c || '').trim());

  // 2. Mapear índice de cada header relevante
  // Buscamos posiciones case-insensitive y tolerante a espacios
  const findIdx = (target, fromIdx = 0) => {
    const t = target.toLowerCase().replace(/\s+/g, ' ').trim();
    for (let i = fromIdx; i < rawHeaders.length; i++) {
      const h = rawHeaders[i].toLowerCase().replace(/\s+/g, ' ').trim();
      if (h === t) return i;
    }
    return -1;
  };

  const idxFecha    = findIdx('fecha vencimiento');
  const idxLinea1   = findIdx('linea de crédito');
  // Buscar SEGUNDA ocurrencia de "Linea de Crédito" (la duplicada)
  const idxLineaDup = idxLinea1 >= 0 ? findIdx('linea de crédito', idxLinea1 + 1) : -1;
  const idxNombre   = findIdx('nombre de cliente');
  const idxCapital  = findIdx('capital');
  const idxInteres  = findIdx('interés');
  const idxOtros    = findIdx('otros');
  const idxIva      = findIdx('iva');
  const idxImporte  = findIdx('importe');
  const idxMoneda   = findIdx('moneda');

  // Validar que las columnas obligatorias se encontraron
  const requiredChecks = [
    ['Fecha Vencimiento', idxFecha],
    ['Linea de Crédito',  idxLinea1],
    ['Nombre de Cliente', idxNombre],
    ['Capital',           idxCapital],
    ['Interés',           idxInteres],
    ['Otros',             idxOtros],
    ['IVA',               idxIva],
    ['Importe',           idxImporte],
    ['Moneda',            idxMoneda]
  ];
  const faltantes = requiredChecks.filter(([n, i]) => i < 0).map(([n]) => n);
  if (faltantes.length > 0) {
    return {
      ok: false,
      error: 'Rep1: faltan columnas requeridas: ' + faltantes.join(', '),
      mismatches: faltantes.map(n => `Columna no encontrada: "${n}"`)
    };
  }

  // 3. Construir filas normalizadas (10 columnas finales)
  const dataRowsRaw = rows.slice(headerRowIdx + 1);
  const dataRows = [];
  for (const row of dataRowsRaw) {
    // Saltar filas vacías
    if (!row || !row.some(c => c !== null && c !== undefined && c !== '')) continue;

    // Saltar fila de "totales" (típicamente tiene texto "TOTAL" en col Capital o similar)
    const fechaCell = row[idxFecha];
    const lineaCell = row[idxLinea1];
    if (!fechaCell || !lineaCell) continue;

    // Saltar si la línea no es numérica (probablemente una fila de subtotal)
    if (typeof lineaCell === 'string' && !/^\d+$/.test(String(lineaCell).trim())) continue;

    dataRows.push([
      fechaCell,
      lineaCell,
      row[idxNombre] || '',
      numOrZero_(row[idxCapital]),
      numOrZero_(row[idxInteres]),
      numOrZero_(row[idxOtros]),
      numOrZero_(row[idxIva]),
      numOrZero_(row[idxImporte]),
      String(row[idxMoneda] || 'MXN').trim()
    ]);
  }

  if (dataRows.length === 0) {
    return { ok: false, error: 'Rep1: no se encontraron filas de datos válidas después del header.' };
  }

  return {
    ok: true,
    allRows: [REP1_HEADERS].concat(dataRows),
    dataRows: dataRows
  };
}

function numOrZero_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  
  // Limpia el string: quita $, espacios y comas
  let limpio = String(v).replace(/[$\s,]/g, '');
  const n = Number(limpio);
  return isNaN(n) ? 0 : n;
}

// ─── HELPERS COMPARTIDOS ───────────────────────────────────────────────────

function getCacheStatus_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return { loaded: false };

  const ts = sh.getRange('B2').getValue();
  const usr = sh.getRange('D2').getValue();
  if (!ts) return { loaded: false };

  const lastRow = sh.getLastRow();
  const dataCount = Math.max(0, lastRow - 4);

  return {
    loaded: true,
    timestamp: (ts instanceof Date) ? ts.toISOString() : String(ts),
    user: usr || '—',
    rowCount: dataCount
  };
}

/**
 * Validador genérico (usado para Rep9).
 */
function validateReport_(rows, expectedHeaders, label) {
  if (!Array.isArray(rows) || rows.length < 2) {
    return { ok: false, error: `${label}: el archivo no contiene datos.` };
  }

  const headers = rows[0].map(h => String(h || '').trim());
  const expected = expectedHeaders.map(h => String(h).trim());

  if (headers.length < expected.length) {
    return {
      ok: false,
      error: `${label}: se esperaban ${expected.length} columnas, se recibieron ${headers.length}.`
    };
  }

  const mismatches = [];
  for (let i = 0; i < expected.length; i++) {
    const a = (headers[i] || '').toLowerCase().replace(/\s+/g, ' ');
    const b = expected[i].toLowerCase().replace(/\s+/g, ' ');
    if (a !== b) {
      mismatches.push(`Col ${i + 1}: esperaba "${expected[i]}", recibió "${headers[i]}"`);
    }
  }

  if (mismatches.length > 0) {
    return {
      ok: false,
      error: `${label}: las columnas no coinciden con el formato esperado.`,
      mismatches: mismatches
    };
  }

  const dataRows = rows.slice(1).filter(r => r.some(c => c !== null && c !== undefined && c !== ''));

  return {
    ok: true,
    label: label,
    rowCount: dataRows.length,
    sample: dataRows.slice(0, 3)
  };
}

/**
 * Escribe los datos al cache, reemplazando lo anterior.
 */
function saveCache_(rows, sheetName, expectedHeaders, label) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    return { ok: false, error: `Hoja "${sheetName}" no encontrada en el maestro.` };
  }

  const userEmail = Session.getActiveUser().getEmail() || 'desconocido';
  const now = new Date();

  // Limpiar datos previos (preservar banner row 1, metadata row 2-3, header row 4)
  const lastRow = sh.getLastRow();
  if (lastRow > 4) {
    sh.getRange(5, 1, lastRow - 4, sh.getMaxColumns()).clearContent();
  }

  // Metadata
  sh.getRange('A2').setValue('Fecha de carga:').setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
  sh.getRange('B2').setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss').setFontFamily('Arial').setFontSize(10);
  sh.getRange('C2').setValue('Cargado por:').setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
  sh.getRange('D2').setValue(userEmail).setFontFamily('Arial').setFontSize(10);

  // Headers en row 4
  sh.getRange(4, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);

  // Datos a partir de row 5
  const dataRows = rows.slice(1).filter(r => r.some(c => c !== null && c !== undefined && c !== ''));
  if (dataRows.length === 0) {
    return { ok: false, error: `${label}: no hay filas con datos para guardar.` };
  }

  const numCols = expectedHeaders.length;
  const normalized = dataRows.map(r => {
    const row = r.slice(0, numCols);
    while (row.length < numCols) row.push('');
    return row;
  });

  sh.getRange(5, 1, normalized.length, numCols).setValues(normalized);

  return {
    ok: true,
    label: label,
    rowsWritten: normalized.length,
    timestamp: now.toISOString(),
    user: userEmail
  };
}

// ─── INICIALIZACIÓN DEL MAESTRO ────────────────────────────────────────────

/**
 * Verifica que todas las hojas necesarias existan en el spreadsheet maestro.
 * Si falta alguna, la crea con la estructura correcta (banner, headers, formato).
 * Solo crea — NUNCA borra contenido existente.
 *
 * Llamar desde el editor cuando:
 *   - Es la primera vez que se configura el sistema
 *   - Se borró alguna hoja por error
 *   - Se quiere validar que todo está bien
 */
function inicializarMaestro() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const reporte = [];

  // Definición de cada hoja con su estructura
  const hojasRequeridas = [
    {
      nombre: SHEETS.CACHE_REP1,
      titulo: 'CACHE — ÚLTIMO REP1 CARGADO (Vencimientos del Periodo)',
      headers: REP1_HEADERS,
      colWidths: { 1: 100, 2: 100, 3: 280, 4: 100, 5: 100, 6: 100, 7: 100, 8: 110, 9: 70 }
    },
    {
      nombre: SHEETS.CACHE_REP9,
      titulo: 'CACHE — ÚLTIMO REP9 CARGADO (Saldos de Cartera)',
      headers: REP9_HEADERS,
      colWidths: null  // 110 default para todas
    },
    {
      nombre: SHEETS.BITACORA,
      titulo: 'BITÁCORA DE ENVÍOS DE AVISOS DE COBRANZA',
      headers: [
        'Timestamp Envío', 'Fecha Vencimiento', 'Tipo Aviso',
        'Línea Crédito', 'Cliente', 'Correos Destino',
        'Total Aviso', 'Status', 'Mensaje / Error'
      ],
      colWidths: { 1: 160, 2: 100, 3: 90, 4: 100, 5: 280, 6: 280, 7: 110, 8: 90, 9: 280 }
    }
  ];

  for (const h of hojasRequeridas) {
    let sh = ss.getSheetByName(h.nombre);
    let creada = false;

    if (!sh) {
      sh = ss.insertSheet(h.nombre);
      creada = true;
    } else {
      // Si ya existe pero tiene contenido raro (ej: solo 1 fila), no la tocamos
      if (sh.getLastRow() >= 4) {
        reporte.push(`✓ ${h.nombre}: ya existe con datos (intacta)`);
        continue;
      }
    }

    // Construir banner row 1
    const numCols = h.headers.length;
    const lastCol = colToLetter_(numCols);
    sh.getRange(1, 1).setValue(h.titulo);
    sh.getRange(`A1:${lastCol}1`).merge()
      .setFontSize(14).setFontWeight('bold').setFontColor('#515151')
      .setBackground('#FDB913')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(1, 28);

    // Row 2: metadata (vacía, se llena al cargar)
    sh.getRange('A2').setValue('Fecha de carga:').setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    sh.getRange('B2').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sh.getRange('C2').setValue('Cargado por:').setFontWeight('bold').setFontFamily('Arial').setFontSize(10);

    // Row 4: headers
    sh.getRange(4, 1, 1, numCols).setValues([h.headers]);
    sh.getRange(4, 1, 1, numCols)
      .setFontWeight('bold')
      .setBackground('#515151')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontFamily('Arial').setFontSize(10);
    sh.setRowHeight(4, 32);

    // Anchos
    if (h.colWidths) {
      for (const [col, width] of Object.entries(h.colWidths)) {
        sh.setColumnWidth(Number(col), width);
      }
    } else {
      for (let c = 1; c <= numCols; c++) sh.setColumnWidth(c, 110);
    }

    sh.setFrozenRows(4);

    reporte.push(creada
      ? `✅ ${h.nombre}: CREADA con estructura correcta`
      : `🔧 ${h.nombre}: existía pero estaba vacía → estructura aplicada`);
  }

  // Verificar que Tasas, Correos y Config existan (estas tienen sus propios datos)
  ['Tasas', 'Correos', 'Config'].forEach(nombre => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) {
      reporte.push(`⚠️ ${nombre}: NO ENCONTRADA — debes restaurarla manualmente`);
    } else {
      const filas = sh.getLastRow();
      reporte.push(`✓ ${nombre}: ok (${filas} filas)`);
    }
  });

  // Mostrar reporte en log y en alert si se ejecuta desde el editor
  const mensaje = reporte.join('\n');
  Logger.log(mensaje);

  try {
    // Si se ejecuta desde el editor con un usuario activo, mostrar UI alert
    SpreadsheetApp.getUi().alert('Inicialización del Maestro', mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // Si se ejecuta desde otro contexto (sin UI), solo log
  }

  return { ok: true, reporte: reporte };
}

// Helper para convertir índice de columna a letra (1=A, 27=AA, etc.)
function colToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
