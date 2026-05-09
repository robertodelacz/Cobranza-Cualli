/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA — Hoja de Trabajo (WorkSheet.gs)
 * ═══════════════════════════════════════════════════════════════════════════
 *  Réplica de la "Hoja 1" del archivo viejo: vuelca el contenido de Cache_Rep1
 *  y agrega las fórmulas VLOOKUP y de cálculo en cols I-O.
 *
 *  Saldos_Vencidos:
 *    Cols A-AD (1-30): copia exacta del Cache_Rep9
 *    Col AE (31): Suma de Moratorios   = R + S + T + U
 *    Col AF (32): Intereses Vencidos    = P + Q
 *
 *  Hoja_Trabajo (estructura):
 *    Row 1: A1='Fecha de Reporte:', B1=TODAY()
 *    Row 2: headers
 *    Row 3+: datos de Rep1 + cálculos
 *
 *    A: Fecha Venc.  B: Línea  C: Cliente  D: Capital  E: Intereses
 *    F: Otros  G: IVA  H: Importe Rep1  I: Moneda
 *    J: Cap. Vencido (VLOOKUP)
 *    K: Intereses Vencidos (VLOOKUP a AF)
 *    L: Suma Moratorios (VLOOKUP a AE)
 *    M: Moratorios del Periodo (calculado)
 *    N: Días (calculado)
 *    O: Tasa Moratoria (×2)
 *    P: TOTAL
 * ═══════════════════════════════════════════════════════════════════════════
 */

const HOJA_TRABAJO = 'Hoja_Trabajo';
const HOJA_SALDOS_VENCIDOS = 'Saldos_Vencidos';

// ─── API PRINCIPAL ─────────────────────────────────────────────────────────

function regenerarHojaTrabajo() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const rep1 = ss.getSheetByName(SHEETS.CACHE_REP1);
    const rep9 = ss.getSheetByName(SHEETS.CACHE_REP9);
    if (!rep1 || rep1.getLastRow() < 5) {
      return { ok: false, error: 'No hay Rep1 cargado.' };
    }
    if (!rep9 || rep9.getLastRow() < 5) {
      return { ok: false, error: 'No hay Rep9 cargado.' };
    }

    const filasSaldos = construirSaldosVencidos_(ss);
    const filasTrabajo = construirHojaTrabajo_(ss);

    return {
      ok: true,
      filasSaldos: filasSaldos,
      filasTrabajo: filasTrabajo,
      timestamp: new Date().toISOString()
    };
  } catch (err) {
    return { ok: false, error: err.message, stack: err.stack };
  }
}

// ─── BUILDERS ──────────────────────────────────────────────────────────────

/**
 * Saldos_Vencidos: copia de Cache_Rep9 (cols A-AD = 30 cols del Rep9)
 * + 2 columnas calculadas:
 *   AE = Suma Moratorios (R+S+T+U)
 *   AF = Intereses Vencidos (P+Q)
 */
function construirSaldosVencidos_(ss) {
  let sh = ss.getSheetByName(HOJA_SALDOS_VENCIDOS);
  if (sh) sh.clear();
  else sh = ss.insertSheet(HOJA_SALDOS_VENCIDOS);

  const rep9 = ss.getSheetByName(SHEETS.CACHE_REP9);
  const lastRowRep9 = rep9.getLastRow();
  const numFilasDatos = lastRowRep9 - 4;

  // Row 1: metadata
  sh.getRange('A1').setValue('Fecha de Reporte:').setFontWeight('bold');
  sh.getRange('B1').setFormula('=TODAY()').setNumberFormat('yyyy-mm-dd');

  // Row 2: headers (30 originales + 2 calculados)
  const rep9Headers = rep9.getRange(4, 1, 1, 30).getValues()[0];
  const headers = rep9Headers.concat(['Suma Moratorios', 'Intereses Vencidos']);
  sh.getRange(2, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 2, headers.length);

  if (numFilasDatos <= 0) return 0;

  // Row 3+: referencias a Cache_Rep9 + cálculos
  // Mapeo de columnas en Saldos_Vencidos (mismo orden que Rep9):
  //   K (11) = Cap. Vencido
  //   P (16) = Int. Vencido
  //   Q (17) = IVA Int. Vencido
  //   R (18) = Moratorios
  //   S (19) = IVA Moratorios
  //   T (20) = Mor. Contabilizados
  //   U (21) = IVA Mor. Contabilizados
  const formulasRef = [];
  for (let r = 0; r < numFilasDatos; r++) {
    const fila = [];
    for (let c = 0; c < 30; c++) {
      const colLetter = colToLetter_(c + 1);
      fila.push(`=IFERROR(Cache_Rep9!${colLetter}${r + 5},"")`);
    }
    const filaActual = r + 3;
    // AE: Suma Moratorios = R + S + T + U
    fila.push(`=IFERROR(R${filaActual}+S${filaActual}+T${filaActual}+U${filaActual},0)`);
    // AF: Intereses Vencidos = P + Q
    fila.push(`=IFERROR(P${filaActual}+Q${filaActual},0)`);
    formulasRef.push(fila);
  }

  sh.getRange(3, 1, numFilasDatos, 32).setFormulas(formulasRef);

  // Formato números (cols numéricas)
  sh.getRange(3, 7, numFilasDatos, 26).setNumberFormat('#,##0.00');

  // Anchos
  for (let c = 1; c <= 32; c++) sh.setColumnWidth(c, 110);
  sh.setColumnWidth(2, 280);

  // Highlight de las 2 columnas calculadas (AE, AF)
  sh.getRange(2, 31, numFilasDatos + 1, 2).setBackground('#FFF8E1');
  sh.getRange(2, 31, 1, 2).setBackground('#FDB913');  // headers en amarillo fuerte

  sh.setFrozenRows(2);
  return numFilasDatos;
}

/**
 * Hoja_Trabajo:
 *   Cols A-I: Cache_Rep1 (9 cols del nuevo Cache)
 *   Cols J-P: cálculos
 */
function construirHojaTrabajo_(ss) {
  let sh = ss.getSheetByName(HOJA_TRABAJO);
  if (sh) sh.clear();
  else sh = ss.insertSheet(HOJA_TRABAJO);

  const rep1 = ss.getSheetByName(SHEETS.CACHE_REP1);
  const lastRowRep1 = rep1.getLastRow();
  const numFilas = lastRowRep1 - 4;

  // Row 1: metadata
  sh.getRange('A1').setValue('Fecha de Reporte:').setFontWeight('bold');
  sh.getRange('B1').setFormula('=TODAY()').setNumberFormat('yyyy-mm-dd');

  // Row 2: headers
  const headers = [
    'Fecha Venc.',           // A
    'Línea',                 // B
    'Cliente',               // C
    'Capital',               // D
    'Intereses',             // E
    'Otros',                 // F
    'IVA',                   // G
    'Importe Rep1',          // H
    'Moneda',                // I
    'Cap. Vencido',          // J — VLOOKUP a Saldos col K (11)
    'Intereses Vencidos',    // K — VLOOKUP a Saldos col AF (32)
    'Suma Moratorios',       // L — VLOOKUP a Saldos col AE (31)
    'Mor. del Periodo',      // M — calculado
    'Días',                  // N — calculado
    'Tasa Moratoria',        // O — VLOOKUP a Tasas × 2
    'TOTAL'                  // P — calculado
  ];
  sh.getRange(2, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 2, headers.length);

  if (numFilas <= 0) return 0;

  // Row 3+: combinación de referencias a Rep1 y fórmulas
  const formulas = [];
  for (let r = 0; r < numFilas; r++) {
    const rRep1 = r + 5;
    const rT = r + 3;

    formulas.push([
      `=IFERROR(Cache_Rep1!A${rRep1},"")`,                                     // A — Fecha Venc
      `=IFERROR(Cache_Rep1!B${rRep1},"")`,                                     // B — Línea
      `=IFERROR(Cache_Rep1!C${rRep1},"")`,                                     // C — Cliente
      `=IFERROR(Cache_Rep1!D${rRep1},0)`,                                      // D — Capital
      `=IFERROR(Cache_Rep1!E${rRep1},0)`,                                      // E — Intereses
      `=IFERROR(Cache_Rep1!F${rRep1},0)`,                                      // F — Otros
      `=IFERROR(Cache_Rep1!G${rRep1},0)`,                                      // G — IVA
      `=IFERROR(Cache_Rep1!H${rRep1},0)`,                                      // H — Importe Rep1
      `=IFERROR(Cache_Rep1!I${rRep1},"MXN")`,                                  // I — Moneda
      // J: Cap. Vencido ← Saldos col K (11)
      `=IFERROR(VLOOKUP(B${rT},Saldos_Vencidos!$A$3:$AF,11,FALSE),0)`,
      // K: Intereses Vencidos ← Saldos col AF (32) = P+Q
      `=IFERROR(VLOOKUP(B${rT},Saldos_Vencidos!$A$3:$AF,32,FALSE),0)`,
      // L: Suma Moratorios ← Saldos col AE (31) = R+S+T+U
      `=IFERROR(VLOOKUP(B${rT},Saldos_Vencidos!$A$3:$AF,31,FALSE),0)`,
      // M: Moratorios del Periodo = J × O / 360 × N (solo si dias > 0 y cap > 0)
      `=IF(AND(J${rT}>0,N${rT}>0),J${rT}*O${rT}/360*N${rT},0)`,
      // N: Días = A − $B$1 (forzando aritmética con +0)
      `=IFERROR((A${rT}+0)-($B$1+0),0)`,
      // O: Tasa Moratoria = Tasas col 3 × 2
      `=IFERROR(VLOOKUP(B${rT},Tasas!$A$3:$C,3,FALSE)*2,0)`,
      // P: TOTAL = D + E + F + G + J + K + L + M
      `=D${rT}+E${rT}+F${rT}+G${rT}+J${rT}+K${rT}+L${rT}+M${rT}`
    ]);
  }

  sh.getRange(3, 1, numFilas, 16).setFormulas(formulas);

  // Formato
  sh.getRange(3, 1, numFilas, 1).setNumberFormat('dd/mm/yyyy');                  // A: fecha
  sh.getRange(3, 4, numFilas, 5).setNumberFormat('#,##0.00');                    // D-H: montos
  sh.getRange(3, 9, numFilas, 1).setHorizontalAlignment('center');               // I: moneda centrada
  sh.getRange(3, 10, numFilas, 4).setNumberFormat('#,##0.00');                   // J-M: montos
  sh.getRange(3, 14, numFilas, 1).setNumberFormat('0');                          // N: días
  sh.getRange(3, 15, numFilas, 1).setNumberFormat('0.00%');                      // O: tasa
  sh.getRange(3, 16, numFilas, 1).setNumberFormat('#,##0.00').setFontWeight('bold');  // P: total

  // Highlight col P (total) y bordes de las cols calculadas
  sh.getRange(3, 16, numFilas, 1).setBackground('#FFF8E1');
  sh.getRange(2, 10, 1, 7).setBackground('#FDB913');  // headers calculados en amarillo

  // Anchos
  sh.setColumnWidth(1, 95);    // Fecha
  sh.setColumnWidth(2, 75);    // Línea
  sh.setColumnWidth(3, 240);   // Cliente
  for (let c = 4; c <= 8; c++) sh.setColumnWidth(c, 105);  // D-H
  sh.setColumnWidth(9, 70);    // Moneda
  for (let c = 10; c <= 16; c++) sh.setColumnWidth(c, 115); // J-P

  sh.setFrozenRows(2);
  sh.setFrozenColumns(3);

  // Fila de totales al final (informativa)
  const filaTotales = numFilas + 4;
  sh.getRange(filaTotales, 1).setValue('TOTALES (informativo)').setFontWeight('bold')
    .setBackground('#515151').setFontColor('#FFFFFF');
  sh.getRange(filaTotales, 1, 1, 16).setBackground('#515151').setFontColor('#FFFFFF');
  // Sumar cols D, E, F, G, H, J, K, L, M, P
  ['D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'P'].forEach(letter => {
    const col = letter.charCodeAt(0) - 64;
    sh.getRange(filaTotales, col).setFormula(
      `=SUM(${letter}3:${letter}${numFilas + 2})`
    ).setNumberFormat('#,##0.00').setFontWeight('bold');
  });

  return numFilas;
}

// ─── HELPERS ───────────────────────────────────────────────────────────────

function formatHeaderRow_(sh, row, numCols) {
  const range = sh.getRange(row, 1, 1, numCols);
  range.setFontWeight('bold')
       .setBackground('#515151')
       .setFontColor('#FFFFFF')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setFontFamily('Arial')
       .setFontSize(10);
  sh.setRowHeight(row, 32);
}

function colToLetter_(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}
