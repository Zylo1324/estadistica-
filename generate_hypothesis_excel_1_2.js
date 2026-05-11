const ExcelJS = require('exceljs');
const path = require('path');

const outputPath = path.join(
  __dirname,
  'TRABAJO DE HIPOTESIS-EJERCICIO1Y2_RESUELTO_CORREGIDO.xlsx'
);

const border = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

const blue = '1F4E79';
const red = 'FF0000';

function normInv(p) {
  // Peter J. Acklam approximation, enough precision for the worksheet values.
  const a = [
    -3.969683028665376e1,
    2.209460984245205e2,
    -2.759285104469687e2,
    1.38357751867269e2,
    -3.066479806614716e1,
    2.506628277459239,
  ];
  const b = [
    -5.447609879822406e1,
    1.615858368580409e2,
    -1.556989798598866e2,
    6.680131188771972e1,
    -1.328068155288572e1,
  ];
  const c = [
    -7.784894002430293e-3,
    -3.223964580411365e-1,
    -2.400758277161838,
    -2.549732539343734,
    4.374664141464968,
    2.938163982698783,
  ];
  const d = [
    7.784695709041462e-3,
    3.224671290700398e-1,
    2.445134137142996,
    3.754408661907416,
  ];

  const plow = 0.02425;
  const phigh = 1 - plow;
  let q;
  let r;

  if (p < plow) {
    q = Math.sqrt(-2 * Math.log(p));
    return (
      (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
      ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
    );
  }

  if (p <= phigh) {
    q = p - 0.5;
    r = q * q;
    return (
      (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) *
      q /
      (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1)
    );
  }

  q = Math.sqrt(-2 * Math.log(1 - p));
  return -(
    (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
    ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1)
  );
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
  const y =
    1 -
    (((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t) *
      Math.exp(-value * value);
  return sign * y;
}

function normCdf(x) {
  return 0.5 * (1 + erf(x / Math.sqrt(2)));
}

function setBaseSheet(ws) {
  ws.views = [{ showGridLines: true }];
  for (let i = 1; i <= 25; i += 1) {
    ws.getColumn(i).width = 10;
  }
  ws.getColumn(1).width = 2;
  ws.getColumn(2).width = 5;
  ws.getColumn(3).width = 14;
  ws.getColumn(4).width = 9;
  ws.getColumn(5).width = 14;
  ws.getColumn(6).width = 3;
  ws.getColumn(7).width = 14;
  ws.getColumn(8).width = 17;
  ws.getColumn(9).width = 14;
  ws.getColumn(10).width = 3;
  ws.getColumn(11).width = 18;
  ws.getColumn(12).width = 8;
  ws.getColumn(13).width = 14;
  ws.getColumn(14).width = 20;

  ws.eachRow((row) => {
    row.font = { name: 'Calibri', size: 11 };
    row.alignment = { vertical: 'middle' };
  });
}

function cell(ws, address, value, opts = {}) {
  const c = ws.getCell(address);
  c.value = value;
  c.font = {
    name: 'Calibri',
    size: opts.size || 11,
    bold: Boolean(opts.bold),
    color: opts.color ? { argb: opts.color } : undefined,
  };
  c.alignment = {
    vertical: 'middle',
    horizontal: opts.align || 'left',
    wrapText: opts.wrap !== false,
  };
  if (opts.fill) {
    c.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: opts.fill },
    };
  }
  if (opts.border) c.border = border;
  if (opts.numFmt) c.numFmt = opts.numFmt;
  return c;
}

function boxed(ws, range) {
  const [from, to] = range.split(':');
  const start = ws.getCell(from);
  const end = ws.getCell(to);
  for (let r = start.row; r <= end.row; r += 1) {
    for (let col = start.col; col <= end.col; col += 1) {
      ws.getCell(r, col).border = border;
    }
  }
}

function merge(ws, range, value, opts = {}) {
  ws.mergeCells(range);
  return cell(ws, range.split(':')[0], value, opts);
}

function richConclusion(ws, address, before, redWord, after) {
  const c = ws.getCell(address);
  c.value = {
    richText: [
      { text: before, font: { name: 'Calibri', size: 11 } },
      { text: redWord, font: { name: 'Calibri', size: 11, color: { argb: red } } },
      { text: after, font: { name: 'Calibri', size: 11 } },
    ],
  };
  c.alignment = { wrapText: true, vertical: 'middle' };
  c.border = border;
}

function format2(value) {
  return Number(value).toFixed(2);
}

function addExerciseSheet(workbook, data) {
  const ws = workbook.addWorksheet(data.sheetName);
  setBaseSheet(ws);

  merge(ws, 'B2:I2', data.title, { bold: true, size: 24, color: red });
  ws.getRow(2).height = 30;
  merge(ws, 'B5:I8', data.statement, { size: 12 });
  ws.getRow(5).height = 25;
  ws.getRow(6).height = 25;
  ws.getRow(7).height = 25;
  ws.getRow(8).height = 25;
  cell(ws, 'B10', 'a)', { bold: true, color: blue, size: 13 });
  merge(ws, 'C10:I11', data.partA, { size: 11 });
  cell(ws, 'B12', 'b)', { bold: true, color: blue, size: 13 });
  merge(ws, 'C12:I13', data.partB, { size: 11 });

  merge(ws, 'B16:E16', 'A) Prueba de hipótesis', { bold: true });
  merge(ws, 'C18:E18', 'Hipótesis', { bold: true, border: true });
  cell(ws, 'D19', 'H0:', { border: true, align: 'center' });
  cell(ws, 'E19', data.h0, { border: true });
  cell(ws, 'D20', 'H1:', { border: true, align: 'center' });
  cell(ws, 'E20', data.h1, { border: true });
  boxed(ws, 'C18:E20');

  merge(ws, 'C23:D23', 'Nivel de significancia:', { bold: true, border: true });
  cell(ws, 'D24', 'α', { border: true, align: 'center' });
  cell(ws, 'E24', data.alpha, { border: true, numFmt: '0.00' });
  boxed(ws, 'C23:E24');

  merge(ws, 'C28:D28', 'Estadístico:', { bold: true, border: true });
  cell(ws, 'D29', 'Zc=', { border: true, align: 'center' });
  cell(ws, 'E29', { formula: data.zcFormula, result: data.zc }, { border: true, numFmt: '0.00000000' });
  boxed(ws, 'C28:E29');

  merge(ws, 'C32:E32', 'DATOS:', { bold: true, border: true });
  cell(ws, 'D33', 'x̄=', { border: true, align: 'center' });
  cell(ws, 'E33', data.xbar, { border: true, numFmt: '0.00' });
  cell(ws, 'D34', 'µ0=', { border: true, align: 'center' });
  cell(ws, 'E34', data.mu0, { border: true, numFmt: '0.00' });
  cell(ws, 'D35', data.sigmaLabel, { border: true, align: 'center' });
  cell(ws, 'E35', { formula: data.sigmaFormula, result: data.sigma }, { border: true, numFmt: '0.00' });
  cell(ws, 'D36', 'n=', { border: true, align: 'center' });
  cell(ws, 'E36', data.n, { border: true, numFmt: '0' });
  boxed(ws, 'C32:E36');

  cell(ws, 'G18', 'Decisión:', { bold: true, border: true });
  cell(ws, 'G19', 'Zt=', { border: true });
  cell(ws, 'H19', { formula: '_xlfn.NORM.S.INV(1-E24)', result: data.zt }, { border: true, numFmt: '0.00000000' });
  merge(ws, 'H20:I20', data.decision, { border: true });
  boxed(ws, 'G18:I20');

  cell(ws, 'G22', 'p-valor=', { bold: true, border: true });
  cell(ws, 'H22', { formula: '1-_xlfn.NORM.S.DIST(E29,1)', result: data.pvalue }, { border: true, numFmt: '0.00000000' });
  merge(ws, 'G24:I25', 'Si p valor es menor o igual que alfa, entonces rechazamos H0', { border: false });

  cell(ws, 'G27', 'Conclusión:', { bold: true, border: true });
  merge(ws, 'H27:I30', '', { border: true });
  richConclusion(ws, 'H27', data.conclusionBefore, data.conclusionRed, data.conclusionAfter);
  boxed(ws, 'G27:I30');

  merge(ws, 'K16:N16', 'B) Intervalo de confianza:', { bold: true });
  cell(ws, 'K18', 'Confianza:', { bold: true, border: true });
  cell(ws, 'L18', 'NC=', { border: true, align: 'center' });
  cell(ws, 'M18', data.confidence, { border: true, numFmt: '0%' });
  cell(ws, 'K19', 'Error:', { bold: true, border: true });
  cell(ws, 'L19', 'E=', { border: true, align: 'center' });
  cell(ws, 'M19', { formula: '_xlfn.NORM.S.INV((1+M18)/2)*E35/SQRT(E36)', result: data.error }, { border: true, numFmt: '0.000000' });
  cell(ws, 'K20', 'Límite Inferior:', { border: true });
  cell(ws, 'L20', 'LI=', { border: true, align: 'center' });
  cell(ws, 'M20', { formula: 'E33-M19', result: data.li }, { border: true, numFmt: '0.000000' });
  cell(ws, 'K21', 'Límite Superior:', { border: true });
  cell(ws, 'L21', 'LS=', { border: true, align: 'center' });
  cell(ws, 'M21', { formula: 'E33+M19', result: data.ls }, { border: true, numFmt: '0.000000' });
  boxed(ws, 'K18:M21');

  cell(ws, 'K25', 'Interpretación:', { bold: true, border: true });
  merge(ws, 'L25:N28', data.interpretation, { border: true, align: 'center' });
  boxed(ws, 'K25:N28');

  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.2,
      footer: 0.2,
    },
  };
}

const exercise1 = (() => {
  const xbar = 600;
  const mu0 = 650;
  const sigma = Math.sqrt(10000);
  const n = 50;
  const alpha = 0.05;
  const confidence = 0.95;
  const zc = (xbar - mu0) / (sigma / Math.sqrt(n));
  const zt = normInv(1 - alpha);
  const pvalue = 1 - normCdf(zc);
  const error = normInv((1 + confidence) / 2) * sigma / Math.sqrt(n);
  const li = xbar - error;
  const ls = xbar + error;
  return {
    sheetName: 'Ejercicio 1',
    title: 'Ejemplo 1',
    statement:
      'El salario básico mensual promedio de los trabajadores no calificados de una empresa sigue una distribución normal con una varianza de $10000. Se ha tomado una muestra aleatoria de 50 trabajadores no calificados y se encontró que tienen un ingreso promedio de $600 mensuales.',
    partA:
      'Compruebe si el salario básico mensual promedio de los trabajadores no calificados de esta empresa es mayor a los $650 mensuales. Utilice una significancia del 5%.',
    partB:
      'Determine un intervalo de 95% de confianza para el verdadero salario básico promedio.',
    h0: 'µ ≤ 650',
    h1: 'µ > 650',
    alpha,
    confidence,
    xbar,
    mu0,
    sigmaLabel: 'σ=',
    sigmaFormula: 'SQRT(10000)',
    sigma,
    n,
    zcFormula: '(E33-E34)/(E35/SQRT(E36))',
    zc,
    zt,
    pvalue,
    decision: pvalue <= alpha ? 'Rechazar H0' : 'No rechazar H0',
    conclusionBefore:
      'El salario básico mensual promedio de los trabajadores no calificados de esta empresa ',
    conclusionRed: 'NO',
    conclusionAfter:
      ' es mayor a los $650 mensuales, con una significancia de 0.05.',
    error,
    li,
    ls,
    interpretation:
      `El verdadero salario básico mensual promedio de los trabajadores no calificados de esta empresa se encuentra entre $${format2(li)} y $${format2(ls)} mensuales, con una confianza de 95%.`,
  };
})();

const exercise2 = (() => {
  const xbar = 11400;
  const mu0 = 11500;
  const sigma = 700;
  const n = 36;
  const alpha = 0.03;
  const confidence = 0.92;
  const zc = (xbar - mu0) / (sigma / Math.sqrt(n));
  const zt = normInv(1 - alpha);
  const pvalue = 1 - normCdf(zc);
  const error = normInv((1 + confidence) / 2) * sigma / Math.sqrt(n);
  const li = xbar - error;
  const ls = xbar + error;
  return {
    sheetName: 'Ejercicio 2',
    title: 'Ejemplo 2',
    statement:
      'El administrador de una empresa generadora de energía que importa carbón, desea estimar un intervalo de confianza de la cantidad de carbón que se consumió por término medio semanalmente durante año pasado. Para ello toma una muestra aleatoria de 36 semanas, resultando que el consumo medio fue de 11400 toneladas, la desviación estándar 700 toneladas. Se conoce que el consumo sigue una distribución aproximadamente normal.',
    partA:
      'Si sus operarios le refieren que el gasto promedio semanal es de más de 11500 toneladas. ¿Qué podría decir acerca de esta afirmación? Use una significancia del 3%',
    partB:
      '¿Cuál será el intervalo de confianza del 92% para el consumo medio semanal durante el año pasado?',
    h0: 'µ ≤ 11500',
    h1: 'µ > 11500',
    alpha,
    confidence,
    xbar,
    mu0,
    sigmaLabel: 's=',
    sigmaFormula: '700',
    sigma,
    n,
    zcFormula: '(E33-E34)/(E35/SQRT(E36))',
    zc,
    zt,
    pvalue,
    decision: pvalue <= alpha ? 'Rechazar H0' : 'No rechazar H0',
    conclusionBefore:
      'El consumo promedio semanal de carbón ',
    conclusionRed: 'NO',
    conclusionAfter:
      ' es mayor a las 11500 toneladas, con una significancia de 0.03.',
    error,
    li,
    ls,
    interpretation:
      `El verdadero consumo medio semanal de carbón se encuentra entre ${format2(li)} y ${format2(ls)} toneladas, con una confianza de 92%.`,
  };
})();

async function main() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Codex';
  workbook.created = new Date();
  addExerciseSheet(workbook, exercise1);
  addExerciseSheet(workbook, exercise2);
  await workbook.xlsx.writeFile(outputPath);
  console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
