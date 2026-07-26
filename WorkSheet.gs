/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  COBRANZA PREVENTIVA — Hoja de Trabajo (WorkSheet.gs) · VERSIÓN MASTER
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
    
    if (!rep1 || rep1.getLastRow() < 5) return { ok: false, error: 'No hay Rep1 cargado.' };
    if (!rep9 || rep9.getLastRow() < 5) return { ok: false, error: 'No hay Rep9 cargado.' };

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

function construirSaldosVencidos_(ss) {
  let sh = ss.getSheetByName(HOJA_SALDOS_VENCIDOS);
  if (sh) sh.clear();
  else sh = ss.insertSheet(HOJA_SALDOS_VENCIDOS);

  const rep9 = ss.getSheetByName(SHEETS.CACHE_REP9);
  const lastRowRep9 = rep9.getLastRow();
  const numFilasDatos = lastRowRep9 - 4;

  // Row 1: metadata
  sh.getRange('A1').setValue('Fecha de Reporte:').setFontWeight('bold');
  sh.getRange('B1').setFormula('=TODAY()').setNumberFormat('dd/MM/yyyy');

  // Row 2: headers
  const rep9Headers = rep9.getRange(4, 1, 1, 30).getValues()[0];
  const headers = rep9Headers.concat(['Suma Moratorios', 'Intereses Vencidos']);
  sh.getRange(2, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 2, headers.length);

  if (numFilasDatos <= 0) return 0;

  const formulasRef = [];
  for (let r = 0; r < numFilasDatos; r++) {
    const fila = [];
    for (let c = 0; c < 30; c++) {
      const colLetter = colToLetter_(c + 1);
      fila.push(`=IF(Cache_Rep9!${colLetter}${r + 5}="","",Cache_Rep9!${colLetter}${r + 5})`);
    }
    const filaActual = r + 3;
    // AE: Suma Moratorios = R + S + T + U
    fila.push(`=IFERROR(R${filaActual}+S${filaActual}+T${filaActual}+U${filaActual},0)`);
    // AF: Intereses Vencidos = P + Q
    fila.push(`=IFERROR(P${filaActual}+Q${filaActual},0)`);
    formulasRef.push(fila);
  }

  sh.getRange(3, 1, numFilasDatos, 32).setFormulas(formulasRef);

  // Formato y anchos
  sh.getRange(3, 5, numFilasDatos, 2).setNumberFormat('dd/MM/yyyy'); // Fechas Rep9
  sh.getRange(3, 7, numFilasDatos, 26).setNumberFormat('#,##0.00');  // Montos
  
  for (let c = 1; c <= 32; c++) sh.setColumnWidth(c, 110);
  sh.setColumnWidth(2, 280); // Cliente

  sh.getRange(2, 31, numFilasDatos + 1, 2).setBackground('#FFF8E1');
  sh.getRange(2, 31, 1, 2).setBackground('#FDB913');

  sh.setFrozenRows(2);
  return numFilasDatos;
}

function construirHojaTrabajo_(ss) {
  let sh = ss.getSheetByName(HOJA_TRABAJO);
  if (sh) sh.clear();
  else sh = ss.insertSheet(HOJA_TRABAJO);

  const rep1 = ss.getSheetByName(SHEETS.CACHE_REP1);
  const lastRowRep1 = rep1.getLastRow();
  const numFilas = lastRowRep1 - 4;

  sh.getRange('A1').setValue('Fecha de Reporte:').setFontWeight('bold');
  sh.getRange('B1').setFormula('=TODAY()').setNumberFormat('dd/MM/yyyy');

  const headers = [
    'Fecha Venc.', 'Línea', 'Cliente', 'Capital', 'Intereses', 'Otros', 'IVA', 'Importe Rep1', 'Moneda',
    'Cap. Vencido', 'Intereses Vencidos', 'Suma Moratorios', 'Mor. del Periodo', 'Días', 'Tasa Moratoria', 'TOTAL'
  ];
  sh.getRange(2, 1, 1, headers.length).setValues([headers]);
  formatHeaderRow_(sh, 2, headers.length);

  if (numFilas <= 0) return 0;

  const formulas = [];
  for (let r = 0; r < numFilas; r++) {
    const rRep1 = r + 5;
    const rT = r + 3;

    formulas.push([
      `=IF(Cache_Rep1!A${rRep1}="","",Cache_Rep1!A${rRep1})`,
      `=IF(Cache_Rep1!B${rRep1}="","",Cache_Rep1!B${rRep1})`,
      `=IF(Cache_Rep1!C${rRep1}="","",Cache_Rep1!C${rRep1})`,
      `=IFERROR(Cache_Rep1!D${rRep1},0)`,
      `=IFERROR(Cache_Rep1!E${rRep1},0)`,
      `=IFERROR(Cache_Rep1!F${rRep1},0)`,
      `=IFERROR(Cache_Rep1!G${rRep1},0)`,
      `=IFERROR(Cache_Rep1!H${rRep1},0)`,
      `=IF(Cache_Rep1!I${rRep1}="","MXN",Cache_Rep1!I${rRep1})`,
      `=IF($B${rT}="","",IFERROR(VLOOKUP($B${rT},Saldos_Vencidos!$A$3:$AF,11,FALSE),0))`,
      `=IF($B${rT}="","",IFERROR(VLOOKUP($B${rT},Saldos_Vencidos!$A$3:$AF,32,FALSE),0))`,
      `=IF($B${rT}="","",IFERROR(VLOOKUP($B${rT},Saldos_Vencidos!$A$3:$AF,31,FALSE),0))`,
      `=IF(AND(J${rT}>0,N${rT}>0),J${rT}*O${rT}/360*N${rT},0)`,
      `=IF(A${rT}="","",IFERROR(A${rT}-$B$1,0))`,
      `=IF($B${rT}="","",IFERROR(VLOOKUP($B${rT},Tasas!$A$3:$C,3,FALSE)*2,0))`,
      `=IF($B${rT}="","",D${rT}+E${rT}+F${rT}+G${rT}+J${rT}+K${rT}+L${rT}+M${rT})`
    ]);
  }

  sh.getRange(3, 1, numFilas, 16).setFormulas(formulas);

  // Formato visual
  sh.getRange(3, 1, numFilas, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(3, 4, numFilas, 5).setNumberFormat('#,##0.00');
  sh.getRange(3, 9, numFilas, 1).setHorizontalAlignment('center');
  sh.getRange(3, 10, numFilas, 4).setNumberFormat('#,##0.00');
  sh.getRange(3, 14, numFilas, 1).setNumberFormat('0');
  sh.getRange(3, 15, numFilas, 1).setNumberFormat('0.00%');
  sh.getRange(3, 16, numFilas, 1).setNumberFormat('#,##0.00').setFontWeight('bold');

  sh.getRange(3, 16, numFilas, 1).setBackground('#FFF8E1');
  sh.getRange(2, 10, 1, 7).setBackground('#FDB913');

  // Dimensiones
  sh.setColumnWidth(1, 95);
  sh.setColumnWidth(2, 75);
  sh.setColumnWidth(3, 240);
  for (let c = 4; c <= 8; c++) sh.setColumnWidth(c, 105);
  sh.setColumnWidth(9, 70);
  for (let c = 10; c <= 16; c++) sh.setColumnWidth(c, 115);

  sh.setFrozenRows(2);
  sh.setFrozenColumns(3);

  // Fila de totales
  const filaTotales = numFilas + 4;
  sh.getRange(filaTotales, 1).setValue('TOTALES (informativo)').setFontWeight('bold')
    .setBackground('#515151').setFontColor('#FFFFFF');
  sh.getRange(filaTotales, 1, 1, 16).setBackground('#515151').setFontColor('#FFFFFF');
  
  ['D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'P'].forEach(letter => {
    const col = letter.charCodeAt(0) - 64;
    sh.getRange(filaTotales, col).setFormula(`=SUM(${letter}3:${letter}${numFilas + 2})`)
      .setNumberFormat('#,##0.00').setFontWeight('bold');
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
