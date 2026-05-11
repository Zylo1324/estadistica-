const ExcelJS = require('exceljs');

const OUTPUT_PATH = 'D:/Estadistica Inferencial/Cuestionario 1 (Tipo A) (2)_resuelto.xlsx';

const COLORS = {
  black: 'FF000000',
  white: 'FFFFFFFF',
  red: 'FFC00000',
  grayFill: 'FFF2F2F2',
  blueFill: 'FFDDEBF7',
  greenFill: 'FFE2F0D9',
  yellowFill: 'FFFFF2CC',
  orangeFill: 'FFFCE4D6',
};

function erf(x) {
  const sign = x < 0 ? -1 : 1;
  const absX = Math.abs(x);
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const t = 1 / (1 + p * absX);
  const y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-(absX ** 2));
  return sign * y;
}

function normalCdf(x) {
  return 0.5 * (1 + erf(x / Math.SQRT2));
}

function exactBinomialRange(n, p, minK, maxK) {
  const q = 1 - p;
  let probability = q ** n;
  let total = 0;

  for (let k = 0; k <= maxK; k += 1) {
    if (k >= minK) {
      total += probability;
    }
    probability *= ((n - k) / (k + 1)) * (p / q);
  }

  return total;
}

function a1(row, col) {
  let current = col;
  let label = '';
  while (current > 0) {
    const remainder = (current - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    current = Math.floor((current - 1) / 26);
  }
  return `${label}${row}`;
}

function setThinBorder(cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: COLORS.black } },
    left: { style: 'thin', color: { argb: COLORS.black } },
    bottom: { style: 'thin', color: { argb: COLORS.black } },
    right: { style: 'thin', color: { argb: COLORS.black } },
  };
}

function fillCell(cell, color) {
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: color },
  };
}

function setFormula(cell, formula, result, numFmt) {
  cell.value = { formula, result };
  if (numFmt) {
    cell.numFmt = numFmt;
  }
}

function applyHeader(cell, color = COLORS.blueFill, fontSize = 11) {
  cell.font = { name: 'Calibri', size: fontSize, bold: true, color: { argb: COLORS.black } };
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  fillCell(cell, color);
  setThinBorder(cell);
}

function applyLabel(ws, row, startCol, endCol, text, fill = COLORS.grayFill) {
  ws.mergeCells(`${a1(row, startCol)}:${a1(row, endCol)}`);
  const cell = ws.getCell(row, startCol);
  cell.value = text;
  cell.font = { name: 'Calibri', size: 11 };
  cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  fillCell(cell, fill);
  for (let col = startCol; col <= endCol; col += 1) {
    setThinBorder(ws.getCell(row, col));
  }
}

function applyValue(ws, row, startCol, endCol, value, numFmt) {
  ws.mergeCells(`${a1(row, startCol)}:${a1(row, endCol)}`);
  const cell = ws.getCell(row, startCol);
  cell.value = value;
  cell.font = { name: 'Calibri', size: 11 };
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  if (numFmt) {
    cell.numFmt = numFmt;
  }
  for (let col = startCol; col <= endCol; col += 1) {
    setThinBorder(ws.getCell(row, col));
  }
}

function applyFormulaValue(ws, row, startCol, endCol, formula, result, numFmt) {
  ws.mergeCells(`${a1(row, startCol)}:${a1(row, endCol)}`);
  const cell = ws.getCell(row, startCol);
  setFormula(cell, formula, result, numFmt);
  cell.font = { name: 'Calibri', size: 11 };
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  for (let col = startCol; col <= endCol; col += 1) {
    setThinBorder(ws.getCell(row, col));
  }
  return cell;
}

function drawBlockFrame(ws, topRow, startCol, endCol, height) {
  for (let row = topRow; row < topRow + height; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      setThinBorder(ws.getCell(row, col));
    }
  }
}

function buildSymmetricMeanBlock(ws, config) {
  const {
    topRow,
    startCol,
    title,
    statement,
    mu,
    sigma,
    n,
    N,
    delta,
    useFpc,
    note,
    resultProbability,
  } = config;
  const endCol = startCol + 5;
  const valueStart = startCol + 4;
  const valueEnd = startCol + 5;
  const nAddr = a1(topRow + 5, valueStart);
  const NAddr = a1(topRow + 4, valueStart);
  const deltaAddr = a1(topRow + 6, valueStart);
  const sigmaAddr = a1(topRow + 3, valueStart);
  const fpcAddr = a1(topRow + 7, valueStart);
  const seAddr = a1(topRow + 8, valueStart);
  const zAddr = a1(topRow + 9, valueStart);

  ws.mergeCells(`${a1(topRow, startCol)}:${a1(topRow, endCol)}`);
  const titleCell = ws.getCell(topRow, startCol);
  titleCell.value = title;
  applyHeader(titleCell, COLORS.blueFill, 12);

  ws.mergeCells(`${a1(topRow + 1, startCol)}:${a1(topRow + 1, endCol)}`);
  const statementCell = ws.getCell(topRow + 1, startCol);
  statementCell.value = statement;
  statementCell.font = { name: 'Calibri', size: 11 };
  statementCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  fillCell(statementCell, COLORS.orangeFill);
  ws.getRow(topRow + 1).height = 42;

  applyLabel(ws, topRow + 2, startCol, startCol + 2, 'Media poblacional');
  applyValue(ws, topRow + 2, valueStart, valueEnd, mu, '0.00');

  applyLabel(ws, topRow + 3, startCol, startCol + 2, 'Desv. estándar poblacional');
  applyValue(ws, topRow + 3, valueStart, valueEnd, sigma, '0.00');

  applyLabel(ws, topRow + 4, startCol, startCol + 2, 'Población total (N)');
  applyValue(ws, topRow + 4, valueStart, valueEnd, N ?? 'No aplica', N ? '0' : undefined);

  applyLabel(ws, topRow + 5, startCol, startCol + 2, 'Tamaño de muestra (n)');
  applyValue(ws, topRow + 5, valueStart, valueEnd, n, '0');

  applyLabel(ws, topRow + 6, startCol, startCol + 2, 'Máxima diferencia permitida');
  applyValue(ws, topRow + 6, valueStart, valueEnd, delta, '0.00');

  applyLabel(ws, topRow + 7, startCol, startCol + 2, 'Factor de corrección finita');
  if (useFpc && N) {
    applyFormulaValue(
      ws,
      topRow + 7,
      valueStart,
      valueEnd,
      `SQRT((${NAddr}-${nAddr})/(${NAddr}-1))`,
      Math.sqrt((N - n) / (N - 1)),
      '0.000000'
    );
  } else {
    applyValue(ws, topRow + 7, valueStart, valueEnd, 1, '0.000000');
  }

  applyLabel(ws, topRow + 8, startCol, startCol + 2, 'Error estándar de X̄');
  applyFormulaValue(
    ws,
    topRow + 8,
    valueStart,
    valueEnd,
    `(${sigmaAddr}/SQRT(${nAddr}))*${fpcAddr}`,
    sigma / Math.sqrt(n) * (useFpc && N ? Math.sqrt((N - n) / (N - 1)) : 1),
    '0.000000'
  );

  applyLabel(ws, topRow + 9, startCol, startCol + 2, 'Valor Z');
  applyFormulaValue(ws, topRow + 9, valueStart, valueEnd, `${deltaAddr}/${seAddr}`, delta / (sigma / Math.sqrt(n) * (useFpc && N ? Math.sqrt((N - n) / (N - 1)) : 1)), '0.000000');

  applyLabel(ws, topRow + 10, startCol, startCol + 2, 'Probabilidad');
  applyFormulaValue(
    ws,
    topRow + 10,
    valueStart,
    valueEnd,
    `(2*NORMSDIST(${zAddr}))-1`,
    resultProbability,
    '0.000000'
  );

  applyLabel(ws, topRow + 11, startCol, startCol + 2, 'Respuesta final', COLORS.yellowFill);
  const answerCell = applyFormulaValue(
    ws,
    topRow + 11,
    valueStart,
    valueEnd,
    a1(topRow + 10, valueStart),
    resultProbability,
    '0.00%'
  );
  answerCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };
  fillCell(answerCell, COLORS.greenFill);

  ws.mergeCells(`${a1(topRow + 12, startCol)}:${a1(topRow + 12, endCol)}`);
  const noteCell = ws.getCell(topRow + 12, startCol);
  noteCell.value = note;
  noteCell.font = { name: 'Calibri', size: 10, italic: true };
  noteCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  fillCell(noteCell, COLORS.white);
  ws.getRow(topRow + 12).height = 30;

  drawBlockFrame(ws, topRow, startCol, endCol, 13);
}

function buildAsymmetricMeanBlock(ws, config) {
  const {
    topRow,
    startCol,
    title,
    statement,
    mu,
    sigma,
    n,
    lower,
    upper,
    probabilityFormula,
    resultProbability,
    note,
  } = config;
  const endCol = startCol + 5;
  const valueStart = startCol + 4;
  const valueEnd = startCol + 5;
  const sigmaAddr = a1(topRow + 3, valueStart);
  const nAddr = a1(topRow + 4, valueStart);
  const seAddr = a1(topRow + 5, valueStart);
  const lowerAddr = a1(topRow + 6, valueStart);
  const upperAddr = a1(topRow + 7, valueStart);
  const muAddr = a1(topRow + 2, valueStart);
  const zLowerAddr = a1(topRow + 8, valueStart);
  const zUpperAddr = a1(topRow + 9, valueStart);

  ws.mergeCells(`${a1(topRow, startCol)}:${a1(topRow, endCol)}`);
  const titleCell = ws.getCell(topRow, startCol);
  titleCell.value = title;
  applyHeader(titleCell, COLORS.blueFill, 12);

  ws.mergeCells(`${a1(topRow + 1, startCol)}:${a1(topRow + 1, endCol)}`);
  const statementCell = ws.getCell(topRow + 1, startCol);
  statementCell.value = statement;
  statementCell.font = { name: 'Calibri', size: 11 };
  statementCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  fillCell(statementCell, COLORS.orangeFill);
  ws.getRow(topRow + 1).height = 42;

  applyLabel(ws, topRow + 2, startCol, startCol + 2, 'Media poblacional');
  applyValue(ws, topRow + 2, valueStart, valueEnd, mu, '0.00');

  applyLabel(ws, topRow + 3, startCol, startCol + 2, 'Desv. estándar poblacional');
  applyValue(ws, topRow + 3, valueStart, valueEnd, sigma, '0.00');

  applyLabel(ws, topRow + 4, startCol, startCol + 2, 'Tamaño de muestra (n)');
  applyValue(ws, topRow + 4, valueStart, valueEnd, n, '0');

  applyLabel(ws, topRow + 5, startCol, startCol + 2, 'Error estándar de X̄');
  applyFormulaValue(ws, topRow + 5, valueStart, valueEnd, `${sigmaAddr}/SQRT(${nAddr})`, sigma / Math.sqrt(n), '0.000000');

  applyLabel(ws, topRow + 6, startCol, startCol + 2, 'Límite inferior');
  applyValue(ws, topRow + 6, valueStart, valueEnd, lower ?? 'No aplica', lower == null ? undefined : '0.00');

  applyLabel(ws, topRow + 7, startCol, startCol + 2, 'Límite superior');
  applyValue(ws, topRow + 7, valueStart, valueEnd, upper ?? 'No aplica', upper == null ? undefined : '0.00');

  applyLabel(ws, topRow + 8, startCol, startCol + 2, 'Z inferior');
  if (lower == null) {
    applyValue(ws, topRow + 8, valueStart, valueEnd, 'No aplica');
  } else {
    applyFormulaValue(
      ws,
      topRow + 8,
      valueStart,
      valueEnd,
      `(${lowerAddr}-${muAddr})/${seAddr}`,
      (lower - mu) / (sigma / Math.sqrt(n)),
      '0.000000'
    );
  }

  applyLabel(ws, topRow + 9, startCol, startCol + 2, 'Z superior');
  if (upper == null) {
    applyValue(ws, topRow + 9, valueStart, valueEnd, 'No aplica');
  } else {
    applyFormulaValue(
      ws,
      topRow + 9,
      valueStart,
      valueEnd,
      `(${upperAddr}-${muAddr})/${seAddr}`,
      (upper - mu) / (sigma / Math.sqrt(n)),
      '0.000000'
    );
  }

  applyLabel(ws, topRow + 10, startCol, startCol + 2, 'Probabilidad');
  applyFormulaValue(ws, topRow + 10, valueStart, valueEnd, probabilityFormula({ zLowerAddr, zUpperAddr }), resultProbability, '0.000000');

  applyLabel(ws, topRow + 11, startCol, startCol + 2, 'Respuesta final', COLORS.yellowFill);
  const answerCell = applyFormulaValue(ws, topRow + 11, valueStart, valueEnd, a1(topRow + 10, valueStart), resultProbability, '0.00%');
  answerCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };
  fillCell(answerCell, COLORS.greenFill);

  ws.mergeCells(`${a1(topRow + 12, startCol)}:${a1(topRow + 12, endCol)}`);
  const noteCell = ws.getCell(topRow + 12, startCol);
  noteCell.value = note;
  noteCell.font = { name: 'Calibri', size: 10, italic: true };
  noteCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  ws.getRow(topRow + 12).height = 28;

  drawBlockFrame(ws, topRow, startCol, endCol, 13);
}

function buildProportionBlock(ws, config) {
  const {
    topRow,
    startCol,
    title,
    statement,
    p,
    n,
    N,
    lower,
    upper,
    resultProbability,
    probabilityFormula,
    note,
    useFpc,
  } = config;
  const endCol = startCol + 5;
  const valueStart = startCol + 4;
  const valueEnd = startCol + 5;
  const pAddr = a1(topRow + 2, valueStart);
  const qAddr = a1(topRow + 3, valueStart);
  const NAddr = a1(topRow + 4, valueStart);
  const nAddr = a1(topRow + 5, valueStart);
  const lowerAddr = a1(topRow + 6, valueStart);
  const upperAddr = a1(topRow + 7, valueStart);
  const fpcAddr = a1(topRow + 8, valueStart);
  const seAddr = a1(topRow + 9, valueStart);
  const zLowerAddr = a1(topRow + 10, valueStart);
  const zUpperAddr = a1(topRow + 11, valueStart);
  const q = 1 - p;
  const fpc = useFpc && N ? Math.sqrt((N - n) / (N - 1)) : 1;
  const se = Math.sqrt((p * q) / n) * fpc;

  ws.mergeCells(`${a1(topRow, startCol)}:${a1(topRow, endCol)}`);
  const titleCell = ws.getCell(topRow, startCol);
  titleCell.value = title;
  applyHeader(titleCell, COLORS.blueFill, 12);

  ws.mergeCells(`${a1(topRow + 1, startCol)}:${a1(topRow + 1, endCol)}`);
  const statementCell = ws.getCell(topRow + 1, startCol);
  statementCell.value = statement;
  statementCell.font = { name: 'Calibri', size: 11 };
  statementCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  fillCell(statementCell, COLORS.orangeFill);
  ws.getRow(topRow + 1).height = 45;

  applyLabel(ws, topRow + 2, startCol, startCol + 2, 'Proporción poblacional (p)');
  applyValue(ws, topRow + 2, valueStart, valueEnd, p, '0.00%');

  applyLabel(ws, topRow + 3, startCol, startCol + 2, 'Complemento (q)');
  applyFormulaValue(ws, topRow + 3, valueStart, valueEnd, `1-${pAddr}`, q, '0.00%');

  applyLabel(ws, topRow + 4, startCol, startCol + 2, 'Población total (N)');
  applyValue(ws, topRow + 4, valueStart, valueEnd, N ?? 'No aplica', N ? '0' : undefined);

  applyLabel(ws, topRow + 5, startCol, startCol + 2, 'Tamaño de muestra (n)');
  applyValue(ws, topRow + 5, valueStart, valueEnd, n, '0');

  applyLabel(ws, topRow + 6, startCol, startCol + 2, 'Límite inferior');
  applyValue(ws, topRow + 6, valueStart, valueEnd, lower ?? 'No aplica', lower == null ? undefined : '0.00%');

  applyLabel(ws, topRow + 7, startCol, startCol + 2, 'Límite superior');
  applyValue(ws, topRow + 7, valueStart, valueEnd, upper ?? 'No aplica', upper == null ? undefined : '0.00%');

  applyLabel(ws, topRow + 8, startCol, startCol + 2, 'Factor de corrección finita');
  if (useFpc && N) {
    applyFormulaValue(ws, topRow + 8, valueStart, valueEnd, `SQRT((${NAddr}-${nAddr})/(${NAddr}-1))`, fpc, '0.000000');
  } else {
    applyValue(ws, topRow + 8, valueStart, valueEnd, 1, '0.000000');
  }

  applyLabel(ws, topRow + 9, startCol, startCol + 2, 'Error estándar de p̂');
  applyFormulaValue(ws, topRow + 9, valueStart, valueEnd, `SQRT((${pAddr}*${qAddr})/${nAddr})*${fpcAddr}`, se, '0.000000');

  applyLabel(ws, topRow + 10, startCol, startCol + 2, 'Z inferior');
  if (lower == null) {
    applyValue(ws, topRow + 10, valueStart, valueEnd, 'No aplica');
  } else {
    applyFormulaValue(ws, topRow + 10, valueStart, valueEnd, `(${lowerAddr}-${pAddr})/${seAddr}`, (lower - p) / se, '0.000000');
  }

  applyLabel(ws, topRow + 11, startCol, startCol + 2, 'Z superior');
  if (upper == null) {
    applyValue(ws, topRow + 11, valueStart, valueEnd, 'No aplica');
  } else {
    applyFormulaValue(ws, topRow + 11, valueStart, valueEnd, `(${upperAddr}-${pAddr})/${seAddr}`, (upper - p) / se, '0.000000');
  }

  applyLabel(ws, topRow + 12, startCol, startCol + 2, 'Probabilidad');
  applyFormulaValue(ws, topRow + 12, valueStart, valueEnd, probabilityFormula({ lowerAddr, upperAddr, pAddr, seAddr, zLowerAddr, zUpperAddr }), resultProbability, '0.000000');

  applyLabel(ws, topRow + 13, startCol, startCol + 2, 'Respuesta final', COLORS.yellowFill);
  const answerCell = applyFormulaValue(ws, topRow + 13, valueStart, valueEnd, a1(topRow + 12, valueStart), resultProbability, '0.00%');
  answerCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };
  fillCell(answerCell, COLORS.greenFill);

  ws.mergeCells(`${a1(topRow + 14, startCol)}:${a1(topRow + 14, endCol)}`);
  const noteCell = ws.getCell(topRow + 14, startCol);
  noteCell.value = note;
  noteCell.font = { name: 'Calibri', size: 10, italic: true };
  noteCell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  ws.getRow(topRow + 14).height = 34;

  drawBlockFrame(ws, topRow, startCol, endCol, 15);
}

function buildSummarySheet(ws, results) {
  ws.columns = [
    { width: 13 },
    { width: 11 },
    { width: 35 },
    { width: 15 },
    { width: 14 },
    { width: 45 },
  ];

  ws.mergeCells('A1:F1');
  const title = ws.getCell('A1');
  title.value = 'Cuestionario 1 (Tipo A) - Resumen de respuestas';
  title.font = { name: 'Calibri', size: 16, bold: true };
  title.alignment = { horizontal: 'center', vertical: 'middle' };

  const headers = ['Pregunta', 'Inciso', 'Distribución', 'Probabilidad', 'Porcentaje', 'Observación'];
  headers.forEach((label, index) => {
    const cell = ws.getCell(3, index + 1);
    cell.value = label;
    applyHeader(cell, COLORS.blueFill);
  });

  const rows = [
    ['1', 'a', 'Media muestral', results.q1a.probability, results.q1a.probability, results.q1a.note],
    ['1', 'b', 'Media muestral', results.q1b.probability, results.q1b.probability, results.q1b.note],
    ['2', 'a', 'Media muestral', results.q2a.probability, results.q2a.probability, results.q2a.note],
    ['2', 'b', 'Media muestral', results.q2b.probability, results.q2b.probability, results.q2b.note],
    ['3', 'a', 'Media muestral', results.q3a.probability, results.q3a.probability, results.q3a.note],
    ['3', 'b', 'Media muestral', results.q3b.probability, results.q3b.probability, results.q3b.note],
    ['4', 'a', 'Proporción muestral', results.q4a.probability, results.q4a.probability, results.q4a.note],
    ['4', 'b', 'Proporción muestral', results.q4b.probability, results.q4b.probability, results.q4b.note],
    ['5', '-', 'Proporción muestral', results.q5.probability, results.q5.probability, results.q5.note],
  ];

  rows.forEach((row, index) => {
    const targetRow = index + 4;
    row.forEach((value, colIndex) => {
      const cell = ws.getCell(targetRow, colIndex + 1);
      cell.value = value;
      cell.font = { name: 'Calibri', size: 11 };
      cell.alignment = {
        vertical: 'middle',
        horizontal: colIndex < 2 ? 'center' : 'left',
        wrapText: true,
      };
      setThinBorder(cell);
      if (colIndex === 3) {
        cell.numFmt = '0.000000';
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
      if (colIndex === 4) {
        cell.numFmt = '0.00%';
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        fillCell(cell, COLORS.greenFill);
      }
    });
  });

  ws.mergeCells('A15:F15');
  const foot = ws.getCell('A15');
  foot.value = 'Criterio usado: corrección por población finita cuando n/N > 5%; en la pregunta 5 se usa la aproximación normal de la distribución muestral de la proporción.';
  foot.font = { name: 'Calibri', size: 10, italic: true };
  foot.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  ws.getRow(15).height = 32;
}

function buildDevelopmentSheet(ws, results) {
  ws.columns = [
    { width: 4 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 10 },
    { width: 12 },
    { width: 12 },
    { width: 15 },
    { width: 15 },
    { width: 15 },
    { width: 10 },
    { width: 12 },
    { width: 12 },
  ];

  ws.mergeCells('B2:M2');
  const title = ws.getCell('B2');
  title.value = 'Evaluación IND 2: Distribución muestral de la media y proporción (Tipo A)';
  title.font = { name: 'Calibri', size: 16, bold: true };
  title.alignment = { horizontal: 'center', vertical: 'middle' };

  ws.mergeCells('B3:M3');
  const subtitle = ws.getCell('B3');
  subtitle.value = 'Desarrollo en Excel con fórmulas visibles y respuestas finales resaltadas.';
  subtitle.font = { name: 'Calibri', size: 11, italic: true };
  subtitle.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(3).height = 22;

  ws.mergeCells('B5:M5');
  const summaryHeader = ws.getCell('B5');
  summaryHeader.value = 'Resumen rápido';
  applyHeader(summaryHeader, COLORS.yellowFill, 12);

  ['Pregunta', 'Inciso', 'Resultado'].forEach((header, index) => {
    const startCol = 2 + index * 4;
    ws.mergeCells(`${a1(6, startCol)}:${a1(6, startCol + 3)}`);
    const cell = ws.getCell(6, startCol);
    cell.value = header;
    applyHeader(cell, COLORS.blueFill);
  });

  const quickRows = [
    ['1', 'a', results.q1a.probability],
    ['1', 'b', results.q1b.probability],
    ['2', 'a', results.q2a.probability],
    ['2', 'b', results.q2b.probability],
    ['3', 'a', results.q3a.probability],
    ['3', 'b', results.q3b.probability],
    ['4', 'a', results.q4a.probability],
    ['4', 'b', results.q4b.probability],
    ['5', '-', results.q5.probability],
  ];

  quickRows.forEach((item, index) => {
    const row = 7 + index;
    ws.mergeCells(`B${row}:E${row}`);
    ws.getCell(`B${row}`).value = `Pregunta ${item[0]}`;
    ws.getCell(`B${row}`).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(`B${row}`).font = { name: 'Calibri', size: 11 };
    fillCell(ws.getCell(`B${row}`), COLORS.grayFill);

    ws.mergeCells(`F${row}:I${row}`);
    ws.getCell(`F${row}`).value = item[1];
    ws.getCell(`F${row}`).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(`F${row}`).font = { name: 'Calibri', size: 11 };

    ws.mergeCells(`J${row}:M${row}`);
    ws.getCell(`J${row}`).value = item[2];
    ws.getCell(`J${row}`).numFmt = '0.00%';
    ws.getCell(`J${row}`).alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getCell(`J${row}`).font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };
    fillCell(ws.getCell(`J${row}`), COLORS.greenFill);

    for (let col = 2; col <= 13; col += 1) {
      setThinBorder(ws.getCell(row, col));
    }
  });

  buildSymmetricMeanBlock(ws, {
    topRow: 18,
    startCol: 2,
    title: 'Pregunta 1.a',
    statement: 'Probabilidad de que el precio medio de 30 gasolineras no difiera en más de S/ 0.05 de la media poblacional.',
    mu: 2.34,
    sigma: 0.20,
    n: 30,
    N: 1000,
    delta: 0.05,
    useFpc: false,
    note: 'No se aplica corrección finita porque n/N = 3%, menor al umbral usual de 5%.',
    resultProbability: results.q1a.probability,
  });

  buildSymmetricMeanBlock(ws, {
    topRow: 18,
    startCol: 8,
    title: 'Pregunta 1.b',
    statement: 'Probabilidad de que el precio medio de 100 gasolineras no difiera en más de S/ 0.05 de la media poblacional.',
    mu: 2.34,
    sigma: 0.20,
    n: 100,
    N: 1000,
    delta: 0.05,
    useFpc: true,
    note: 'Sí se aplica corrección finita porque n/N = 10%, mayor al 5%.',
    resultProbability: results.q1b.probability,
  });

  buildSymmetricMeanBlock(ws, {
    topRow: 33,
    startCol: 2,
    title: 'Pregunta 2.a',
    statement: 'Probabilidad de que la media muestral de 60 personas no se aleje más o menos de 1 hora del promedio semanal.',
    mu: 15,
    sigma: 4,
    n: 60,
    N: null,
    delta: 1,
    useFpc: false,
    note: 'No se usa corrección finita porque no se proporciona N.',
    resultProbability: results.q2a.probability,
  });

  buildSymmetricMeanBlock(ws, {
    topRow: 33,
    startCol: 8,
    title: 'Pregunta 2.b',
    statement: 'Probabilidad de que la media muestral de 60 personas no se aleje más o menos de 45 minutos del promedio semanal.',
    mu: 15,
    sigma: 4,
    n: 60,
    N: null,
    delta: 0.75,
    useFpc: false,
    note: '45 minutos equivalen a 0.75 horas.',
    resultProbability: results.q2b.probability,
  });

  buildAsymmetricMeanBlock(ws, {
    topRow: 48,
    startCol: 2,
    title: 'Pregunta 3.a',
    statement: 'Probabilidad de que la media muestral de 50 facturas sea superior a $1300.',
    mu: 1210,
    sigma: 250,
    n: 50,
    lower: 1300,
    upper: null,
    probabilityFormula: ({ zLowerAddr }) => `1-NORMSDIST(${zLowerAddr})`,
    resultProbability: results.q3a.probability,
    note: 'Se calcula una cola superior: P(X̄ > 1300).',
  });

  buildAsymmetricMeanBlock(ws, {
    topRow: 48,
    startCol: 8,
    title: 'Pregunta 3.b',
    statement: 'Probabilidad de que la media muestral de 50 facturas esté entre $1150 y $1200.',
    mu: 1210,
    sigma: 250,
    n: 50,
    lower: 1150,
    upper: 1200,
    probabilityFormula: ({ zLowerAddr, zUpperAddr }) => `NORMSDIST(${zUpperAddr})-NORMSDIST(${zLowerAddr})`,
    resultProbability: results.q3b.probability,
    note: 'Se calcula una probabilidad entre dos límites.',
  });

  buildProportionBlock(ws, {
    topRow: 63,
    startCol: 2,
    title: 'Pregunta 4.a',
    statement: 'Probabilidad de que la proporción muestral de 326 personas que usan gafas esté entre 20% y 30%.',
    p: 0.25,
    n: 326,
    N: 5000,
    lower: 0.20,
    upper: 0.30,
    useFpc: true,
    probabilityFormula: ({ zLowerAddr, zUpperAddr }) => `NORMSDIST(${zUpperAddr})-NORMSDIST(${zLowerAddr})`,
    resultProbability: results.q4a.probability,
    note: 'Se aplica corrección finita porque n/N = 6.52%, mayor al 5%.',
  });

  buildProportionBlock(ws, {
    topRow: 63,
    startCol: 8,
    title: 'Pregunta 4.b',
    statement: 'Probabilidad de que la proporción muestral de 326 personas que usan gafas sea mayor que 23%.',
    p: 0.25,
    n: 326,
    N: 5000,
    lower: 0.23,
    upper: null,
    useFpc: true,
    probabilityFormula: ({ zLowerAddr }) => `1-NORMSDIST(${zLowerAddr})`,
    resultProbability: results.q4b.probability,
    note: 'Se calcula una cola superior para p̂ > 0.23.',
  });

  buildProportionBlock(ws, {
    topRow: 80,
    startCol: 2,
    title: 'Pregunta 5',
    statement: 'Probabilidad de que en una caja de 200 piezas la proporción defectuosa esté entre 5% y 7%, usando la distribución muestral de la proporción.',
    p: 0.03,
    n: 200,
    N: null,
    lower: 0.05,
    upper: 0.07,
    useFpc: false,
    probabilityFormula: ({ zLowerAddr, zUpperAddr }) => `NORMSDIST(${zUpperAddr})-NORMSDIST(${zLowerAddr})`,
    resultProbability: results.q5.probability,
    note: `Respuesta principal con aproximación normal de p̂. Como contraste exacto binomial: ${results.q5.exactBinomial.toFixed(6)} (${(results.q5.exactBinomial * 100).toFixed(2)}%).`,
  });

  ws.mergeCells('B96:M96');
  const assumptionsHeader = ws.getCell('B96');
  assumptionsHeader.value = 'Notas metodológicas';
  applyHeader(assumptionsHeader, COLORS.yellowFill, 12);

  const notes = [
    '1. Para las preguntas 1.b y 4 se aplicó corrección por población finita porque la fracción muestral supera el 5%.',
    '2. Para las preguntas 2 y 3 no se aplica corrección finita porque el tamaño de la población no fue proporcionado.',
    '3. En la pregunta 5 se dejó como respuesta principal la aproximación normal porque el tema es distribución muestral de la proporción.',
  ];

  notes.forEach((text, index) => {
    const row = 97 + index;
    ws.mergeCells(`${a1(row, 2)}:${a1(row, 13)}`);
    const cell = ws.getCell(row, 2);
    cell.value = text;
    cell.font = { name: 'Calibri', size: 10, italic: true };
    cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    ws.getRow(row).height = 24;
  });
}

async function main() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'OpenAI Codex';
  workbook.created = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;

  const q1a = {
    se: 0.20 / Math.sqrt(30),
  };
  q1a.z = 0.05 / q1a.se;
  q1a.probability = (2 * normalCdf(q1a.z)) - 1;
  q1a.note = 'Sin corrección finita (n/N = 3%).';

  const fpc1b = Math.sqrt((1000 - 100) / (1000 - 1));
  const q1b = {
    se: (0.20 / Math.sqrt(100)) * fpc1b,
  };
  q1b.z = 0.05 / q1b.se;
  q1b.probability = (2 * normalCdf(q1b.z)) - 1;
  q1b.note = 'Con corrección finita (n/N = 10%).';

  const q2a = {
    se: 4 / Math.sqrt(60),
  };
  q2a.z = 1 / q2a.se;
  q2a.probability = (2 * normalCdf(q2a.z)) - 1;
  q2a.note = 'Intervalo simétrico de ±1 hora.';

  const q2b = {
    se: 4 / Math.sqrt(60),
  };
  q2b.z = 0.75 / q2b.se;
  q2b.probability = (2 * normalCdf(q2b.z)) - 1;
  q2b.note = '45 minutos = 0.75 horas.';

  const se3 = 250 / Math.sqrt(50);
  const z3a = (1300 - 1210) / se3;
  const q3a = {
    probability: 1 - normalCdf(z3a),
    note: 'Cola superior.',
  };

  const z3bLower = (1150 - 1210) / se3;
  const z3bUpper = (1200 - 1210) / se3;
  const q3b = {
    probability: normalCdf(z3bUpper) - normalCdf(z3bLower),
    note: 'Probabilidad entre dos valores.',
  };

  const fpc4 = Math.sqrt((5000 - 326) / (5000 - 1));
  const se4 = Math.sqrt((0.25 * 0.75) / 326) * fpc4;
  const q4a = {
    probability: normalCdf((0.30 - 0.25) / se4) - normalCdf((0.20 - 0.25) / se4),
    note: 'Con corrección finita.',
  };
  const q4b = {
    probability: 1 - normalCdf((0.23 - 0.25) / se4),
    note: 'Cola superior con corrección finita.',
  };

  const se5 = Math.sqrt((0.03 * 0.97) / 200);
  const q5 = {
    probability: normalCdf((0.07 - 0.03) / se5) - normalCdf((0.05 - 0.03) / se5),
    exactBinomial: exactBinomialRange(200, 0.03, 10, 14),
    note: 'Aproximación normal de p̂; se deja comparación exacta en el desarrollo.',
  };

  const results = { q1a, q1b, q2a, q2b, q3a, q3b, q4a, q4b, q5 };

  const summarySheet = workbook.addWorksheet('Resumen');
  const developmentSheet = workbook.addWorksheet('Desarrollo');

  buildSummarySheet(summarySheet, results);
  buildDevelopmentSheet(developmentSheet, results);

  await workbook.xlsx.writeFile(OUTPUT_PATH);
  console.log(`Archivo generado en: ${OUTPUT_PATH}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
