const ExcelJS = require('exceljs');
const path = require('path');

const outputPath = path.join(
  __dirname,
  'TRABAJO_DE_HIPOTESIS_EJERCICIO_1_Y_2_FORMATO_MODELO.xlsx'
);

const colors = {
  title: 'FF0000',
  blue: 'DDEBF7',
  peach: 'FCE4D6',
  beige: 'FFF2CC',
  gray: 'EDEDED',
  green: 'E2F0D9',
  white: 'FFFFFF',
  black: '000000',
};

const border = {
  top: { style: 'thin', color: { argb: colors.black } },
  left: { style: 'thin', color: { argb: colors.black } },
  bottom: { style: 'thin', color: { argb: colors.black } },
  right: { style: 'thin', color: { argb: colors.black } },
};

function normInv(p) {
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  const plow = 0.02425;
  const phigh = 1 - plow;
  let q;
  let r;

  if (p < plow) {
    q = Math.sqrt(-2 * Math.log(p));
    return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }

  if (p <= phigh) {
    q = p - 0.5;
    r = q * q;
    return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q /
      (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  }

  q = Math.sqrt(-2 * Math.log(1 - p));
  return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
    ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
}

function erf(x) {
  const sign = x >= 0 ? 1 : -1;
  const value = Math.abs(x);
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  const t = 1 / (1 + p * value);
  const y = 1 - (((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t) * Math.exp(-value * value);
  return sign * y;
}

function normCdf(x) {
  return 0.5 * (1 + erf(x / Math.sqrt(2)));
}

function format2(value) {
  return Number(value).toFixed(2);
}

function setCell(ws, row, col, value, opts = {}) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.font = {
    name: 'Calibri',
    size: opts.size || 11,
    bold: Boolean(opts.bold),
    italic: Boolean(opts.italic),
    color: opts.color ? { argb: opts.color } : undefined,
  };
  cell.alignment = {
    horizontal: opts.align || 'left',
    vertical: 'middle',
    wrapText: opts.wrap !== false,
  };
  if (opts.fill) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: opts.fill },
    };
  }
  if (opts.border !== false) cell.border = border;
  if (opts.numFmt) cell.numFmt = opts.numFmt;
  return cell;
}

function merge(ws, row1, col1, row2, col2, value, opts = {}) {
  ws.mergeCells(row1, col1, row2, col2);
  return setCell(ws, row1, col1, value, opts);
}

function applyBox(ws, row1, col1, row2, col2) {
  for (let row = row1; row <= row2; row += 1) {
    for (let col = col1; col <= col2; col += 1) {
      ws.getCell(row, col).border = border;
    }
  }
}

function setRichNo(ws, row, col, before, after, opts = {}) {
  const cell = ws.getCell(row, col);
  cell.value = {
    richText: [
      { text: before, font: { name: 'Calibri', size: opts.size || 11 } },
      { text: 'NO', font: { name: 'Calibri', size: opts.size || 11, color: { argb: colors.title } } },
      { text: after, font: { name: 'Calibri', size: opts.size || 11 } },
    ],
  };
  cell.alignment = { horizontal: opts.align || 'center', vertical: 'middle', wrapText: true };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.fill || colors.green } };
  cell.border = border;
}

function styleSheet(ws) {
  ws.views = [{ showGridLines: true, state: 'frozen', ySplit: 4 }];
  ws.properties.defaultRowHeight = 19;
  const widths = [3, 5, 12, 12, 13, 3, 12, 14, 14, 3, 14, 10, 14, 24];
  widths.forEach((width, index) => {
    ws.getColumn(index + 1).width = width;
  });
  for (let col = 15; col <= 18; col += 1) ws.getColumn(col).width = 10;
}

function addTitle(ws) {
  merge(ws, 2, 2, 2, 14, 'Trabajo de hipótesis: Estimación y prueba de hipótesis de una media', {
    bold: true,
    size: 14,
    align: 'center',
    border: false,
  });
  merge(ws, 3, 2, 3, 14, 'Desarrollo en Excel con fórmulas y respuestas finales resaltadas.', {
    italic: true,
    align: 'center',
    border: false,
  });
}

function addSummary(ws, exercises) {
  merge(ws, 5, 2, 5, 14, 'Resumen rápido', {
    bold: true,
    align: 'center',
    fill: colors.beige,
  });
  setCell(ws, 6, 2, 'Ejercicio', { bold: true, align: 'center', fill: colors.blue });
  merge(ws, 6, 3, 6, 5, 'Inciso', { bold: true, align: 'center', fill: colors.blue });
  merge(ws, 6, 6, 6, 14, 'Resultado', { bold: true, align: 'center', fill: colors.blue });

  const rows = [
    ['Ejercicio 1', 'a', exercises.e1.testResult],
    ['Ejercicio 1', 'b', exercises.e1.intervalResult],
    ['Ejercicio 2', 'a', exercises.e2.testResult],
    ['Ejercicio 2', 'b', exercises.e2.intervalResult],
  ];

  rows.forEach((item, index) => {
    const row = 7 + index;
    setCell(ws, row, 2, item[0], { align: 'center', fill: colors.gray });
    merge(ws, row, 3, row, 5, item[1], { align: 'center' });
    merge(ws, row, 6, row, 14, item[2], {
      align: 'center',
      fill: colors.green,
      bold: true,
      color: colors.title,
    });
  });
}

function addTestBlock(ws, startRow, startCol, title, statement, cfg) {
  const endCol = startCol + 5;
  merge(ws, startRow, startCol, startRow, endCol, title, {
    bold: true,
    align: 'center',
    fill: colors.blue,
  });
  merge(ws, startRow + 1, startCol, startRow + 2, endCol, statement, {
    fill: colors.peach,
    wrap: true,
  });
  ws.getRow(startRow + 1).height = 34;
  ws.getRow(startRow + 2).height = 34;

  const rows = [
    ['Hipótesis nula', 'H0:', cfg.h0],
    ['Hipótesis alternativa', 'H1:', cfg.h1],
    ['Nivel de significancia', 'α', cfg.alpha],
    ['Media muestral', 'x̄', cfg.xbar],
    ['Media hipotética', 'µ0', cfg.mu0],
    ['Desv. estándar', cfg.sigmaSymbol, { formula: cfg.sigmaFormula, result: cfg.sigma }],
    ['Tamaño de muestra', 'n', cfg.n],
    ['Valor Z calculado', 'Zc', { formula: `${cfg.xbarCell}-${cfg.mu0Cell}`, result: cfg.xbar - cfg.mu0 }],
  ];

  rows.forEach((item, index) => {
    const row = startRow + 3 + index;
    merge(ws, row, startCol, row, startCol + 2, item[0], { fill: colors.gray });
    setCell(ws, row, startCol + 3, item[1], { align: 'center' });
    setCell(ws, row, startCol + 4, item[2], { align: 'right', numFmt: typeof item[2] === 'number' ? '0.00' : undefined });
    setCell(ws, row, startCol + 5, '', {});
  });

  const xbarAddress = ws.getCell(startRow + 6, startCol + 4).address;
  const mu0Address = ws.getCell(startRow + 7, startCol + 4).address;
  const sigmaCell = ws.getCell(startRow + 8, startCol + 4).address;
  const nCell = ws.getCell(startRow + 9, startCol + 4).address;
  const zcCell = ws.getCell(startRow + 10, startCol + 4);
  zcCell.value = { formula: `(${xbarAddress}-${mu0Address})/(${sigmaCell}/SQRT(${nCell}))`, result: cfg.zc };
  zcCell.numFmt = '0.000000';

  const alphaCell = ws.getCell(startRow + 5, startCol + 4).address;
  const zcAddress = zcCell.address;
  const ztRow = startRow + 11;
  merge(ws, ztRow, startCol, ztRow, startCol + 2, 'Valor crítico', { fill: colors.gray });
  setCell(ws, ztRow, startCol + 3, 'Zt', { align: 'center' });
  setCell(ws, ztRow, startCol + 4, { formula: `_xlfn.NORM.S.INV(1-${alphaCell})`, result: cfg.zt }, { align: 'right', numFmt: '0.000000' });
  setCell(ws, ztRow, startCol + 5, '', {});

  const pRow = startRow + 12;
  merge(ws, pRow, startCol, pRow, startCol + 2, 'p-valor', { fill: colors.gray });
  setCell(ws, pRow, startCol + 3, 'p', { align: 'center' });
  setCell(ws, pRow, startCol + 4, { formula: `1-_xlfn.NORM.S.DIST(${zcAddress},1)`, result: cfg.pvalue }, { align: 'right', numFmt: '0.000000' });
  setCell(ws, pRow, startCol + 5, '', {});

  const decisionRow = startRow + 13;
  merge(ws, decisionRow, startCol, decisionRow, startCol + 2, 'Decisión', { fill: colors.beige });
  merge(ws, decisionRow, startCol + 3, decisionRow, endCol, cfg.decision, {
    align: 'center',
    fill: colors.green,
    bold: true,
    color: colors.title,
  });

  const finalRow = startRow + 14;
  merge(ws, finalRow, startCol, finalRow, startCol + 2, 'Respuesta final', { fill: colors.beige });
  merge(ws, finalRow, startCol + 3, finalRow + 1, endCol, '', { fill: colors.green });
  setRichNo(ws, finalRow, startCol + 3, cfg.finalBefore, cfg.finalAfter, { fill: colors.green });

  applyBox(ws, startRow, startCol, finalRow + 1, endCol);
}

function addIntervalBlock(ws, startRow, startCol, title, statement, cfg) {
  const endCol = startCol + 5;
  merge(ws, startRow, startCol, startRow, endCol, title, {
    bold: true,
    align: 'center',
    fill: colors.blue,
  });
  merge(ws, startRow + 1, startCol, startRow + 2, endCol, statement, {
    fill: colors.peach,
    wrap: true,
  });
  ws.getRow(startRow + 1).height = 34;
  ws.getRow(startRow + 2).height = 34;

  const rows = [
    ['Nivel de confianza', 'NC', cfg.confidence],
    ['Media muestral', 'x̄', cfg.xbar],
    ['Desv. estándar', cfg.sigmaSymbol, cfg.sigma],
    ['Tamaño de muestra', 'n', cfg.n],
    ['Valor Z', 'Z', { formula: `_xlfn.NORM.S.INV((1+${ws.getCell(startRow + 3, startCol + 4).address})/2)`, result: cfg.zConfidence }],
  ];

  rows.forEach((item, index) => {
    const row = startRow + 3 + index;
    merge(ws, row, startCol, row, startCol + 2, item[0], { fill: colors.gray });
    setCell(ws, row, startCol + 3, item[1], { align: 'center' });
    setCell(ws, row, startCol + 4, item[2], {
      align: 'right',
      numFmt: item[0] === 'Nivel de confianza' ? '0%' : '0.00',
    });
    setCell(ws, row, startCol + 5, '', {});
  });

  const ncCell = ws.getCell(startRow + 3, startCol + 4).address;
  const xbarCell = ws.getCell(startRow + 4, startCol + 4).address;
  const sigmaCell = ws.getCell(startRow + 5, startCol + 4).address;
  const nCell = ws.getCell(startRow + 6, startCol + 4).address;
  const zCell = ws.getCell(startRow + 7, startCol + 4);
  zCell.value = { formula: `_xlfn.NORM.S.INV((1+${ncCell})/2)`, result: cfg.zConfidence };
  zCell.numFmt = '0.000000';

  const errorRow = startRow + 8;
  merge(ws, errorRow, startCol, errorRow, startCol + 2, 'Error', { fill: colors.gray });
  setCell(ws, errorRow, startCol + 3, 'E', { align: 'center' });
  setCell(ws, errorRow, startCol + 4, { formula: `${zCell.address}*${sigmaCell}/SQRT(${nCell})`, result: cfg.error }, { align: 'right', numFmt: '0.000000' });
  setCell(ws, errorRow, startCol + 5, '', {});

  const liRow = startRow + 9;
  const errorCell = ws.getCell(errorRow, startCol + 4).address;
  merge(ws, liRow, startCol, liRow, startCol + 2, 'Límite inferior', { fill: colors.gray });
  setCell(ws, liRow, startCol + 3, 'LI', { align: 'center' });
  setCell(ws, liRow, startCol + 4, { formula: `${xbarCell}-${errorCell}`, result: cfg.li }, { align: 'right', numFmt: '0.000000' });
  setCell(ws, liRow, startCol + 5, '', {});

  const lsRow = startRow + 10;
  merge(ws, lsRow, startCol, lsRow, startCol + 2, 'Límite superior', { fill: colors.gray });
  setCell(ws, lsRow, startCol + 3, 'LS', { align: 'center' });
  setCell(ws, lsRow, startCol + 4, { formula: `${xbarCell}+${errorCell}`, result: cfg.ls }, { align: 'right', numFmt: '0.000000' });
  setCell(ws, lsRow, startCol + 5, '', {});

  const finalRow = startRow + 11;
  merge(ws, finalRow, startCol, finalRow, startCol + 2, 'Respuesta final', { fill: colors.beige });
  merge(ws, finalRow, startCol + 3, finalRow + 2, endCol, cfg.final, {
    fill: colors.green,
    align: 'center',
    bold: true,
    color: colors.title,
  });
  applyBox(ws, startRow, startCol, finalRow + 2, endCol);
}

function buildData() {
  const e1 = {
    xbar: 600,
    mu0: 650,
    sigma: 100,
    sigmaFormula: 'SQRT(10000)',
    sigmaSymbol: 'σ',
    n: 50,
    alpha: 0.05,
    confidence: 0.95,
  };
  e1.zc = (e1.xbar - e1.mu0) / (e1.sigma / Math.sqrt(e1.n));
  e1.zt = normInv(1 - e1.alpha);
  e1.pvalue = 1 - normCdf(e1.zc);
  e1.zConfidence = normInv((1 + e1.confidence) / 2);
  e1.error = e1.zConfidence * e1.sigma / Math.sqrt(e1.n);
  e1.li = e1.xbar - e1.error;
  e1.ls = e1.xbar + e1.error;
  e1.decision = 'No rechazar H0';
  e1.testResult = 'No se rechaza H0: no es mayor a $650';
  e1.intervalResult = `IC 95%: $${format2(e1.li)} a $${format2(e1.ls)}`;

  const e2 = {
    xbar: 11400,
    mu0: 11500,
    sigma: 700,
    sigmaFormula: '700',
    sigmaSymbol: 's',
    n: 36,
    alpha: 0.03,
    confidence: 0.92,
  };
  e2.zc = (e2.xbar - e2.mu0) / (e2.sigma / Math.sqrt(e2.n));
  e2.zt = normInv(1 - e2.alpha);
  e2.pvalue = 1 - normCdf(e2.zc);
  e2.zConfidence = normInv((1 + e2.confidence) / 2);
  e2.error = e2.zConfidence * e2.sigma / Math.sqrt(e2.n);
  e2.li = e2.xbar - e2.error;
  e2.ls = e2.xbar + e2.error;
  e2.decision = 'No rechazar H0';
  e2.testResult = 'No se rechaza H0: no es mayor a 11500 toneladas';
  e2.intervalResult = `IC 92%: ${format2(e2.li)} a ${format2(e2.ls)} toneladas`;

  return { e1, e2 };
}

function addDevelopment(workbook) {
  const ws = workbook.addWorksheet('Desarrollo');
  styleSheet(ws);
  addTitle(ws);
  const data = buildData();
  addSummary(ws, data);

  addTestBlock(
    ws,
    14,
    2,
    'Ejercicio 1.a',
    'Compruebe si el salario básico mensual promedio de los trabajadores no calificados de esta empresa es mayor a los $650 mensuales. Use α = 5%.',
    {
      ...data.e1,
      h0: 'μ ≤ 650',
      h1: 'μ > 650',
      xbarCell: 'E21',
      mu0Cell: 'E22',
      finalBefore: 'El salario básico mensual promedio ',
      finalAfter: ' es mayor a $650 mensuales.',
    }
  );

  addIntervalBlock(
    ws,
    14,
    9,
    'Ejercicio 1.b',
    'Determine un intervalo de 95% de confianza para el verdadero salario básico promedio.',
    {
      ...data.e1,
      final: `El verdadero salario básico mensual promedio se encuentra entre $${format2(data.e1.li)} y $${format2(data.e1.ls)} mensuales, con una confianza de 95%.`,
    }
  );

  addTestBlock(
    ws,
    34,
    2,
    'Ejercicio 2.a',
    'Si sus operarios refieren que el gasto promedio semanal es de más de 11500 toneladas, compruebe esta afirmación con α = 3%.',
    {
      ...data.e2,
      h0: 'μ ≤ 11500',
      h1: 'μ > 11500',
      xbarCell: 'E41',
      mu0Cell: 'E42',
      finalBefore: 'El consumo promedio semanal de carbón ',
      finalAfter: ' es mayor a 11500 toneladas.',
    }
  );

  addIntervalBlock(
    ws,
    34,
    9,
    'Ejercicio 2.b',
    'Calcule el intervalo de confianza del 92% para el consumo medio semanal durante el año pasado.',
    {
      ...data.e2,
      final: `El verdadero consumo medio semanal de carbón se encuentra entre ${format2(data.e2.li)} y ${format2(data.e2.ls)} toneladas, con una confianza de 92%.`,
    }
  );

  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 2,
    margins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
  };
}

function addSummarySheet(workbook) {
  const ws = workbook.addWorksheet('Resumen');
  styleSheet(ws);
  addTitle(ws);
  addSummary(ws, buildData());
}

async function main() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Codex';
  workbook.created = new Date();
  workbook.views = [{ activeTab: 1 }];
  addSummarySheet(workbook);
  addDevelopment(workbook);
  await workbook.xlsx.writeFile(outputPath);
  console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
