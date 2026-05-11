const ExcelJS = require('exceljs');
const path = require('path');

const inputPath = path.join(
  __dirname,
  'UNIDAD 2',
  'Estimacion_y_prueba_de_hipotesis_media_12_ejercicios_resuelto_CORREGIDO.xlsx'
);

const outputPath = path.join(
  __dirname,
  'UNIDAD 2',
  'esto.xlsx'
);

const red = 'FF0000';
const black = '000000';

const border = {
  top: { style: 'thin', color: { argb: black } },
  left: { style: 'thin', color: { argb: black } },
  bottom: { style: 'thin', color: { argb: black } },
  right: { style: 'thin', color: { argb: black } },
};

function textValue(value) {
  if (value == null) return '';
  if (typeof value === 'object' && value.richText) return value.richText.map((part) => part.text).join('');
  if (typeof value === 'object' && value.formula) return value.result ?? '';
  return value;
}

function plainText(value) {
  const text = textValue(value);
  return text == null ? '' : String(text);
}

function numericValue(cell) {
  const value = cell.value;
  if (typeof value === 'object' && value?.formula) return value.result;
  return value;
}

function resultOrValue(ws, address) {
  return numericValue(ws.getCell(address));
}

function findRow(ws, startRow, endRow, col, needle) {
  for (let row = startRow; row <= endRow; row += 1) {
    if (plainText(ws.getCell(row, col).value).toLowerCase().includes(needle.toLowerCase())) {
      return row;
    }
  }
  return undefined;
}

function inferTail(h1) {
  if (h1.includes('≠') || h1.includes('<>')) return 'two';
  if (h1.includes('<')) return 'left';
  return 'right';
}

function percentText(confidence) {
  const value = Number(confidence) * 100;
  return Number.isInteger(value) ? `${value}%` : `${value.toFixed(1)}%`;
}

function formatNumber(value, decimals = 2) {
  return Number(value).toFixed(decimals);
}

function readSourceWorkbook(workbook) {
  return workbook.worksheets
    .filter((ws) => /^Ejercicio\s+\d+$/i.test(ws.name))
    .map((ws) => {
      const statRow = findRow(ws, 15, 35, 2, 'Estadístico');
      const criticalRow = findRow(ws, 15, 35, 2, 'Valor crítico');
      const pRow = findRow(ws, 15, 35, 2, 'p-valor');
      const decisionRow = findRow(ws, 15, 35, 2, 'Decisión');
      const responseRow = findRow(ws, 15, 35, 2, 'Respuesta final');
      const dfRow = findRow(ws, 15, 35, 2, 'Grados de libertad');
      const method = plainText(ws.getCell(statRow, 5).value).toLowerCase().includes('t') ? 't' : 'z';
      const h1 = plainText(ws.getCell(18, 6).value);

      return {
        number: Number(ws.name.replace(/\D+/g, '')),
        sheetName: ws.name,
        statement: plainText(ws.getCell('B4').value),
        partA: plainText(ws.getCell('C9').value),
        partB: plainText(ws.getCell('C11').value),
        h0: plainText(ws.getCell('F17').value),
        h1,
        alpha: resultOrValue(ws, 'F19'),
        xbar: resultOrValue(ws, 'F20'),
        mu0: resultOrValue(ws, 'F21'),
        sigma: resultOrValue(ws, 'F22'),
        sigmaSymbol: plainText(ws.getCell('E22').value) || (method === 't' ? 's' : 'σ'),
        n: resultOrValue(ws, 'F23'),
        df: dfRow ? resultOrValue(ws, `F${dfRow}`) : undefined,
        method,
        tail: inferTail(h1),
        stat: resultOrValue(ws, `F${statRow}`),
        critical: resultOrValue(ws, `F${criticalRow}`),
        pvalue: resultOrValue(ws, `F${pRow}`),
        decision: plainText(ws.getCell(decisionRow, 5).value),
        finalTest: plainText(ws.getCell(responseRow, 5).value),
        intervals: readIntervals(ws),
        sampleSize: readSampleSize(ws),
      };
    })
    .sort((a, b) => a.number - b.number);
}

function readIntervals(ws) {
  const intervals = [];
  for (let row = 15; row <= 65; row += 1) {
    const title = plainText(ws.getCell(row, 9).value);
    if (!title.toLowerCase().includes('intervalo de confianza')) continue;

    const nextTitleRow = findNextRightBlock(ws, row + 1);
    const endRow = nextTitleRow ? nextTitleRow - 1 : row + 15;
    const confidenceRow = findRow(ws, row + 1, endRow, 9, 'Nivel de confianza');
    const criticalRow = findRow(ws, row + 1, endRow, 9, 'Valor');
    const errorRow = findRow(ws, row + 1, endRow, 9, 'Error');
    const liRow = findRow(ws, row + 1, endRow, 9, 'Límite inferior') || findRow(ws, row + 1, endRow, 9, 'Limite inferior');
    const lsRow = findRow(ws, row + 1, endRow, 9, 'Límite superior') || findRow(ws, row + 1, endRow, 9, 'Limite superior');
    const responseRow = findRow(ws, row + 1, endRow + 4, 9, 'Respuesta final');

    intervals.push({
      title,
      confidence: resultOrValue(ws, `M${confidenceRow}`),
      critical: criticalRow ? resultOrValue(ws, `M${criticalRow}`) : undefined,
      error: resultOrValue(ws, `M${errorRow}`),
      li: resultOrValue(ws, `M${liRow}`),
      ls: resultOrValue(ws, `M${lsRow}`),
      finalText: responseRow ? plainText(ws.getCell(responseRow, 12).value) : '',
    });
  }
  return intervals;
}

function findNextRightBlock(ws, startRow) {
  for (let row = startRow; row <= 70; row += 1) {
    const text = plainText(ws.getCell(row, 9).value).toLowerCase();
    if (text.includes('intervalo de confianza')) return row;
    if (/ejercicio\s+\d+\.d/i.test(text) && text.includes('tamaño de muestra')) return row;
  }
  return undefined;
}

function readSampleSize(ws) {
  let titleRow;
  for (let row = 15; row <= 70; row += 1) {
    const text = plainText(ws.getCell(row, 9).value);
    if (/ejercicio\s+\d+\.d/i.test(text) && text.toLowerCase().includes('tamaño de muestra')) {
      titleRow = row;
      break;
    }
  }
  if (!titleRow) return undefined;

  const rowFor = (label) => findRow(ws, titleRow + 1, titleRow + 12, 9, label);
  const responseRow = rowFor('Respuesta final');
  return {
    title: plainText(ws.getCell(titleRow, 9).value),
    confidence: resultOrValue(ws, `M${rowFor('Nivel de confianza')}`),
    z: resultOrValue(ws, `M${rowFor('Valor Z')}`),
    sigma: resultOrValue(ws, `M${rowFor('Desv')}`),
    error: resultOrValue(ws, `M${rowFor('Margen')}`),
    calculated: resultOrValue(ws, `M${rowFor('Tamaño calculado')}`),
    minimum: resultOrValue(ws, `M${rowFor('Tamaño mínimo')}`),
    finalText: responseRow ? plainText(ws.getCell(responseRow, 12).value) : '',
  };
}

function setBaseSheet(ws) {
  ws.views = [{ showGridLines: true }];
  ws.properties.defaultRowHeight = 16;
  const widths = {
    A: 4,
    B: 11,
    C: 12,
    D: 10,
    E: 15,
    F: 4,
    G: 12,
    H: 12,
    I: 12,
    J: 12,
    K: 12,
    L: 12,
    M: 12,
    N: 4,
    O: 4,
    P: 16,
    Q: 10,
    R: 15,
    S: 8,
    T: 14,
    U: 12,
    V: 12,
    W: 12,
    X: 12,
    Y: 12,
  };
  Object.entries(widths).forEach(([col, width]) => {
    ws.getColumn(col).width = width;
  });
  [1, 2, 3, 4].forEach((row) => {
    ws.getRow(row).height = 20;
  });
  [6, 7, 8, 9].forEach((row) => {
    ws.getRow(row).height = 18;
  });
  ws.getRow(12).height = 20;
}

function styleCell(cell, opts = {}) {
  cell.font = {
    name: 'Calibri',
    size: opts.size || 11,
    bold: Boolean(opts.bold),
    color: opts.color ? { argb: opts.color } : undefined,
  };
  cell.alignment = {
    horizontal: opts.align || 'left',
    vertical: 'middle',
    wrapText: opts.wrap !== false,
  };
  if (opts.border) cell.border = border;
  if (opts.numFmt) cell.numFmt = opts.numFmt;
  return cell;
}

function cell(ws, address, value, opts = {}) {
  const target = ws.getCell(address);
  target.value = value;
  return styleCell(target, opts);
}

function merge(ws, range, value, opts = {}) {
  ws.mergeCells(range);
  const address = range.split(':')[0];
  cell(ws, address, value, opts);
  if (opts.border) boxed(ws, range);
  return ws.getCell(address);
}

function boxed(ws, range) {
  const [from, to] = range.split(':');
  const start = ws.getCell(from);
  const end = ws.getCell(to || from);
  for (let row = start.row; row <= end.row; row += 1) {
    for (let col = start.col; col <= end.col; col += 1) {
      ws.getCell(row, col).border = border;
    }
  }
}

function formula(result, expression) {
  return result;
}

function addRichConclusion(ws, address, text) {
  const target = ws.getCell(address);
  const firstMatch = text.match(/\b(NO|No|SÍ|Sí|Se rechaza|Rechazar H0|No rechazar H0)\b/);
  if (!firstMatch) {
    target.value = text;
  } else {
    const index = firstMatch.index;
    const match = firstMatch[0];
    target.value = {
      richText: [
        { text: text.slice(0, index), font: { name: 'Calibri', size: 11 } },
        { text: match, font: { name: 'Calibri', size: 11, bold: true, color: { argb: red } } },
        { text: text.slice(index + match.length), font: { name: 'Calibri', size: 11 } },
      ],
    };
  }
  target.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  target.border = border;
}

function criticalFormula(ex) {
  if (ex.method === 'z') {
    if (ex.tail === 'left') return '_xlfn.NORM.S.INV(C19)';
    if (ex.tail === 'two') return '_xlfn.NORM.S.INV(1-C19/2)';
    return '_xlfn.NORM.S.INV(1-C19)';
  }
  if (ex.tail === 'left') return '_xlfn.T.INV(C19,C32)';
  if (ex.tail === 'right') return '_xlfn.T.INV(1-C19,C32)';
  return 'TINV(C19,C32)';
}

function pValueFormula(ex) {
  if (ex.method === 'z') {
    if (ex.tail === 'left') return '_xlfn.NORM.S.DIST(C22,1)';
    if (ex.tail === 'two') return '2*(1-_xlfn.NORM.S.DIST(ABS(C22),1))';
    return '1-_xlfn.NORM.S.DIST(C22,1)';
  }
  if (ex.tail === 'left') return '_xlfn.T.DIST(C22,C32,TRUE)';
  if (ex.tail === 'right') return '1-_xlfn.T.DIST(C22,C32,TRUE)';
  return 'TDIST(ABS(C22),C32,2)';
}

function intervalErrorFormula(ex, confidenceCell) {
  if (ex.method === 'z') return `_xlfn.NORM.S.INV((1+${confidenceCell})/2)*C30/SQRT(C31)`;
  return `TINV(1-${confidenceCell},C32)*C30/SQRT(C31)`;
}

function addTop(ws, ex) {
  merge(ws, 'A1:M4', ex.statement, { size: 12, wrap: true });
  cell(ws, 'A6', 'a)', { bold: true, size: 11 });
  merge(ws, 'B6:M7', ex.partA, { size: 11, wrap: true });
  cell(ws, 'A8', 'b)', { bold: true, size: 11 });
  merge(ws, 'B8:M9', ex.partB, { size: 11, wrap: true });
}

function addHypothesisBlock(ws, ex) {
  merge(ws, 'B12:M12', 'A) Prueba de hipotesis', { bold: true, align: 'center' });

  merge(ws, 'B14:E14', 'Hipotesis', { bold: true, border: true, align: 'center' });
  cell(ws, 'B15', 'H0:', { border: true, align: 'center' });
  merge(ws, 'C15:E15', ex.h0, { border: true, align: 'center' });
  cell(ws, 'B16', 'H1:', { border: true, align: 'center' });
  merge(ws, 'C16:E16', ex.h1, { border: true, align: 'center' });
  boxed(ws, 'B14:E16');

  merge(ws, 'B18:D18', 'Nivel de significacia:', { border: true });
  cell(ws, 'B19', 'α', { border: true, align: 'center' });
  merge(ws, 'C19:E19', ex.alpha, { border: true, numFmt: '0.00', align: 'center' });
  boxed(ws, 'B18:E19');

  merge(ws, 'G14:M14', 'Decisión', { bold: true, border: true, align: 'center' });
  const statLetter = ex.method.toUpperCase();
  const criticalLabel = ex.tail === 'two' ? `±${statLetter}t=` : `${statLetter}t=`;
  cell(ws, 'G15', criticalLabel, { border: true });
  merge(ws, 'H15:J15', formula(ex.critical, criticalFormula(ex)), { border: true, numFmt: '0.00000000' });
  merge(ws, 'K15:M15', ex.decision, {
    border: true,
    align: 'center',
  });
  boxed(ws, 'G14:M15');

  cell(ws, 'G17', 'p-valor=', { bold: true, border: true });
  merge(ws, 'H17:J17', formula(ex.pvalue, pValueFormula(ex)), { border: true, numFmt: '0.00000000' });
  merge(ws, 'G19:M20', 'Si p valor es menor o igual que alfa, entonces rechazamos H0', { wrap: true });

  cell(ws, 'G22', 'Conclusión:', { bold: true, border: true });
  merge(ws, 'H22:M25', '', { border: true });
  addRichConclusion(ws, 'H22', ex.finalTest);
  boxed(ws, 'G22:M25');

  merge(ws, 'B21:E21', 'Estadistico:', { border: true });
  const statLabel = `${statLetter}c=`;
  cell(ws, 'B22', statLabel, { border: true, align: 'center' });
  merge(ws, 'C22:E22', formula(ex.stat, '(C28-C29)/(C30/SQRT(C31))'), {
    border: true,
    numFmt: '0.00000000',
    align: 'center',
  });
  boxed(ws, 'B21:E22');

  merge(ws, 'B26:E26', 'DATOS:', { border: true, bold: true });
  cell(ws, 'B28', 'x̄=', { border: true, align: 'center' });
  merge(ws, 'C28:E28', ex.xbar, { border: true, numFmt: '0.00', align: 'center' });
  cell(ws, 'B29', 'µ0=', { border: true, align: 'center' });
  merge(ws, 'C29:E29', ex.mu0, { border: true, numFmt: '0.00', align: 'center' });
  cell(ws, 'B30', `${ex.sigmaSymbol}=`, { border: true, align: 'center' });
  merge(ws, 'C30:E30', ex.sigma, { border: true, numFmt: '0.00', align: 'center' });
  cell(ws, 'B31', 'n=', { border: true, align: 'center' });
  merge(ws, 'C31:E31', ex.n, { border: true, numFmt: '0', align: 'center' });
  const dataRange = ex.method === 't' ? 'B26:E32' : 'B26:E31';
  if (ex.method === 't') {
    cell(ws, 'B32', 'gl=', { border: true, align: 'center' });
    merge(ws, 'C32:E32', ex.df, { border: true, numFmt: '0', align: 'center' });
  }
  boxed(ws, dataRange);
}

function addIntervalBlock(ws, ex, interval, startRow, index) {
  const title = index === 0 ? 'B) Intervalo de confianza:' : interval.title.replace('Intervalo de confianza', 'intervalo de confianza');
  merge(ws, `P${startRow}:Y${startRow}`, title, { bold: true, align: 'center' });
  const confRow = startRow + 2;
  const errorRow = confRow + 1;
  const liRow = confRow + 2;
  const lsRow = confRow + 3;
  const interpretationRow = confRow + 7;

  cell(ws, `P${confRow}`, 'Confianza:', { bold: true, border: true });
  cell(ws, `Q${confRow}`, 'NC=', { border: true, align: 'center' });
  cell(ws, `R${confRow}`, interval.confidence, { border: true, numFmt: '0.0%', align: 'center' });

  cell(ws, `P${errorRow}`, 'Error:', { bold: true, border: true });
  cell(ws, `Q${errorRow}`, 'E=', { border: true, align: 'center' });
  cell(ws, `R${errorRow}`, formula(interval.error, intervalErrorFormula(ex, `R${confRow}`)), {
    border: true,
    numFmt: '0.000000',
    align: 'center',
  });

  cell(ws, `P${liRow}`, 'Limite Inferior:', { border: true });
  cell(ws, `Q${liRow}`, 'LI=', { border: true, align: 'center' });
  cell(ws, `R${liRow}`, formula(interval.li, `C28-R${errorRow}`), { border: true, numFmt: '0.000000', align: 'center' });

  cell(ws, `P${lsRow}`, 'Limite Superior:', { border: true });
  cell(ws, `Q${lsRow}`, 'LS=', { border: true, align: 'center' });
  cell(ws, `R${lsRow}`, formula(interval.ls, `C28+R${errorRow}`), { border: true, numFmt: '0.000000', align: 'center' });
  boxed(ws, `P${confRow}:R${lsRow}`);

  merge(ws, `P${interpretationRow}:R${interpretationRow + 3}`, 'Interpretación:', { bold: true, border: true, align: 'center' });
  merge(ws, `S${interpretationRow}:Y${interpretationRow + 3}`, interval.finalText, {
    border: true,
    align: 'center',
    wrap: true,
  });
  boxed(ws, `P${interpretationRow}:Y${interpretationRow + 3}`);
  return interpretationRow + 5;
}

function addSampleSizeBlock(ws, sample, startRow) {
  if (!sample) return;
  merge(ws, `P${startRow}:Y${startRow}`, sample.title, { bold: true, align: 'center' });
  const rows = [
    ['Confianza:', 'NC=', sample.confidence, '0.0%'],
    ['Valor Z:', 'Z=', formula(sample.z, `_xlfn.NORM.S.INV((1+R${startRow + 1})/2)`), '0.000000'],
    ['Desv. estándar:', 'σ=', sample.sigma, '0.00'],
    ['Margen de error:', 'E=', sample.error, '0.00'],
    ['Tamaño calculado:', 'n=', formula(sample.calculated, `(R${startRow + 2}*R${startRow + 3}/R${startRow + 4})^2`), '0.000000'],
    ['Tamaño mínimo:', 'n=', sample.minimum, '0'],
  ];

  rows.forEach(([label, symbol, value, numFmt], index) => {
    const row = startRow + 1 + index;
    cell(ws, `P${row}`, label, { border: true });
    cell(ws, `Q${row}`, symbol, { border: true, align: 'center' });
    cell(ws, `R${row}`, value, { border: true, numFmt, align: 'center' });
  });
  boxed(ws, `P${startRow + 1}:R${startRow + rows.length}`);

  const responseRow = startRow + rows.length + 2;
  merge(ws, `P${responseRow}:R${responseRow + 2}`, 'Interpretación:', { bold: true, border: true, align: 'center' });
  merge(ws, `S${responseRow}:Y${responseRow + 2}`, sample.finalText, { border: true, align: 'center', wrap: true });
  boxed(ws, `P${responseRow}:Y${responseRow + 2}`);
}

function addExerciseSheet(workbook, ex) {
  const ws = workbook.addWorksheet(ex.sheetName);
  setBaseSheet(ws);
  addTop(ws, ex);
  addHypothesisBlock(ws, ex);

  let nextBlockRow = 12;
  ex.intervals.forEach((interval, index) => {
    nextBlockRow = addIntervalBlock(ws, ex, interval, nextBlockRow, index);
  });
  addSampleSizeBlock(ws, ex.sampleSize, nextBlockRow);

  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { left: 0.2, right: 0.2, top: 0.35, bottom: 0.35, header: 0.2, footer: 0.2 },
  };
}

async function main() {
  const sourceWorkbook = new ExcelJS.Workbook();
  await sourceWorkbook.xlsx.readFile(inputPath);
  const exercises = readSourceWorkbook(sourceWorkbook);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Codex';
  workbook.created = new Date();
  exercises.forEach((exercise) => addExerciseSheet(workbook, exercise));
  await workbook.xlsx.writeFile(outputPath);

  console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
