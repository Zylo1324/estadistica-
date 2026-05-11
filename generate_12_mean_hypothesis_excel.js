const ExcelJS = require('exceljs');
const path = require('path');

const outputPath = path.join(
  __dirname,
  'UNIDAD 2',
  'Estimacion_y_prueba_de_hipotesis_media_12_ejercicios_resuelto_CORREGIDO.xlsx'
);

const colors = {
  red: 'FF0000',
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

function gammaLn(xx) {
  const cof = [
    76.18009172947146,
    -86.50532032941677,
    24.01409824083091,
    -1.231739572450155,
    0.1208650973866179e-2,
    -0.5395239384953e-5,
  ];
  let x = xx - 1;
  let tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  let ser = 1.000000000190015;
  for (let j = 0; j < cof.length; j += 1) {
    x += 1;
    ser += cof[j] / x;
  }
  return -tmp + Math.log(2.5066282746310005 * ser);
}

function betaCf(a, b, x) {
  const maxIterations = 100;
  const eps = 3e-14;
  const fpMin = 1e-300;
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < fpMin) d = fpMin;
  d = 1 / d;
  let h = d;

  for (let m = 1; m <= maxIterations; m += 1) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < fpMin) d = fpMin;
    c = 1 + aa / c;
    if (Math.abs(c) < fpMin) c = fpMin;
    d = 1 / d;
    h *= d * c;

    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d;
    if (Math.abs(d) < fpMin) d = fpMin;
    c = 1 + aa / c;
    if (Math.abs(c) < fpMin) c = fpMin;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < eps) break;
  }
  return h;
}

function incompleteBeta(x, a, b) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(gammaLn(a + b) - gammaLn(a) - gammaLn(b) + a * Math.log(x) + b * Math.log(1 - x));
  if (x < (a + 1) / (a + b + 2)) return (bt * betaCf(a, b, x)) / a;
  return 1 - (bt * betaCf(b, a, 1 - x)) / b;
}

function tCdf(t, df) {
  const x = df / (df + t * t);
  const ib = incompleteBeta(x, df / 2, 0.5);
  return t >= 0 ? 1 - 0.5 * ib : 0.5 * ib;
}

function tInv(p, df) {
  let lo = -20;
  let hi = 20;
  for (let i = 0; i < 200; i += 1) {
    const mid = (lo + hi) / 2;
    if (tCdf(mid, df) < p) lo = mid;
    else hi = mid;
  }
  return (lo + hi) / 2;
}

function mean(values) {
  return values.reduce((acc, value) => acc + value, 0) / values.length;
}

function sampleSd(values) {
  const xbar = mean(values);
  const variance = values.reduce((acc, value) => acc + (value - xbar) ** 2, 0) / (values.length - 1);
  return Math.sqrt(variance);
}

function fmt(value, decimals = 2) {
  return Number(value).toFixed(decimals);
}

function calcExercise(raw) {
  const xbar = raw.values ? mean(raw.values) : raw.xbar;
  const sigma = raw.values ? sampleSd(raw.values) : raw.sigma;
  const n = raw.values ? raw.values.length : raw.n;
  const df = raw.method === 't' ? n - 1 : undefined;
  const stat = (xbar - raw.mu0) / (sigma / Math.sqrt(n));

  let critical;
  let pvalue;
  if (raw.method === 'z') {
    if (raw.tail === 'right') {
      critical = normInv(1 - raw.alpha);
      pvalue = 1 - normCdf(stat);
    } else if (raw.tail === 'left') {
      critical = normInv(raw.alpha);
      pvalue = normCdf(stat);
    } else {
      critical = normInv(1 - raw.alpha / 2);
      pvalue = 2 * (1 - normCdf(Math.abs(stat)));
    }
  } else {
    critical = tInv(1 - raw.alpha / 2, df);
    pvalue = 2 * (1 - tCdf(Math.abs(stat), df));
  }

  const reject =
    raw.tail === 'right'
      ? stat > critical
      : raw.tail === 'left'
        ? stat < critical
        : Math.abs(stat) > critical;

  const interval = raw.intervals.map((intervalSpec) => {
    const crit =
      raw.method === 'z'
        ? normInv((1 + intervalSpec.confidence) / 2)
        : tInv((1 + intervalSpec.confidence) / 2, df);
    const error = crit * sigma / Math.sqrt(n);
    return {
      ...intervalSpec,
      critical: crit,
      error,
      li: xbar - error,
      ls: xbar + error,
    };
  });

  return {
    ...raw,
    xbar,
    sigma,
    n,
    df,
    stat,
    critical,
    pvalue,
    reject,
    decision: reject ? 'Rechazar H0' : 'No rechazar H0',
    intervals: interval,
  };
}

function makeData() {
  const base = [
    {
      number: 1,
      title: 'Ejercicio 1',
      statement:
        'El gerente de la cadena de una tienda afirma que en promedio los clientes gastaron $400 el año pasado. Sin embargo, analizando el mercado, nosotros creemos, que dicho gerente ha exagerado. Para someter a prueba estas hipótesis se tomó una muestra aleatoria de 100 clientes que el año pasado habían comprado en dicha tienda, ésta reveló una media de $450 y una desviación estándar $100.',
      partA: 'En el nivel de significación de 0.05, ¿Es posible concluir que los clientes de esta tienda están gastando más?',
      partB: 'Calcule e interprete su intervalo de confianza al 95%.',
      method: 'z',
      tail: 'right',
      h0: 'μ ≤ 400',
      h1: 'μ > 400',
      alpha: 0.05,
      xbar: 450,
      mu0: 400,
      sigma: 100,
      sigmaLabel: 's',
      n: 100,
      intervals: [{ label: 'b', confidence: 0.95 }],
      finalTest: 'Sí se concluye que los clientes están gastando más de $400.',
      finalInterval: (interval) => `El gasto promedio se encuentra entre $${fmt(interval.li)} y $${fmt(interval.ls)} con 95% de confianza.`,
    },
    {
      number: 2,
      title: 'Ejercicio 2',
      statement:
        'El gerente de ventas de una empresa que elabora cápsulas de uña de gato indica que la demanda semanal tiene distribución normal con una media de 1000 cápsulas y una desviación estándar de 360 cápsulas. Sin embargo, en un estudio reciente, una muestra aleatoria de 36 semanas dio una demanda promedio de 950 cápsulas.',
      partA: '¿Es posible concluir que la producción promedio semanal es menos de 1000 cápsulas al 1% de significación?',
      partB: 'Calcule e interprete su intervalo de confianza al 92%.',
      method: 'z',
      tail: 'left',
      h0: 'μ ≥ 1000',
      h1: 'μ < 1000',
      alpha: 0.01,
      xbar: 950,
      mu0: 1000,
      sigma: 360,
      sigmaLabel: 'σ',
      n: 36,
      intervals: [{ label: 'b', confidence: 0.92 }],
      finalTest: 'No se concluye que la producción promedio semanal sea menor de 1000 cápsulas.',
      finalInterval: (interval) => `La demanda promedio semanal se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} cápsulas con 92% de confianza.`,
    },
    {
      number: 3,
      title: 'Ejercicio 3',
      statement:
        'Un fabricante de aparatos de TV afirma que se necesita a lo mucho 250 micro amperes de corriente para alcanzar cierto grado de brillantez con un tipo de televisor en particular. Una muestra de 20 aparatos de TV produce un promedio muestral de corriente de 257.3 micro amperes. Denotemos por μ el verdadero promedio de corriente necesaria para alcanzar la brillantez deseada con aparatos de este tipo, y supongamos que μ es la media de una población con σ = 15.',
      partA: 'Pruebe al nivel de significación del 5% la hipótesis nula de que la media es a lo sumo 250 micro amperes.',
      partB: 'Calcule e interprete su intervalo de confianza al 90%.',
      method: 'z',
      tail: 'right',
      h0: 'μ ≤ 250',
      h1: 'μ > 250',
      alpha: 0.05,
      xbar: 257.3,
      mu0: 250,
      sigma: 15,
      sigmaLabel: 'σ',
      n: 20,
      intervals: [{ label: 'b', confidence: 0.9 }],
      finalTest: 'Se rechaza que la media sea a lo sumo 250 micro amperes.',
      finalInterval: (interval) => `La corriente promedio necesaria se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} micro amperes con 90% de confianza.`,
    },
    {
      number: 4,
      title: 'Ejercicio 4',
      statement:
        'En una conocida prueba de autoimagen que da por resultado puntajes que se distribuyen de manera normal, se espera que el puntaje medio de los beneficiarios de la asistencia pública sea de 65. La prueba se aplica a una muestra aleatoria de 28 beneficiarios de la asistencia pública quienes logran un puntaje medio de 62.1 con una desviación estándar de 5.83.',
      partA: 'Al nivel de significación del 1%, ¿los beneficiarios de la asistencia pública tienen puntajes diferentes en promedio, a los que se esperaban?',
      partB: 'Calcule e interprete su intervalo de confianza 97%.',
      method: 't',
      tail: 'two',
      h0: 'μ = 65',
      h1: 'μ ≠ 65',
      alpha: 0.01,
      xbar: 62.1,
      mu0: 65,
      sigma: 5.83,
      sigmaLabel: 's',
      n: 28,
      intervals: [{ label: 'b', confidence: 0.97 }],
      finalTest: 'No se demuestra diferencia significativa al 1%.',
      finalInterval: (interval) => `El puntaje promedio se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} puntos con 97% de confianza.`,
    },
    {
      number: 5,
      title: 'Ejercicio 5',
      statement:
        'Para tratar de estimar el consumo promedio por cliente, en un gran restaurante, se reunieron datos de una muestra de 49 clientes durante un periodo de tres semanas. Si la media de la muestra es de $22.60 dólares con una desviación estándar de $2.5.',
      partA: '¿Existe evidencia para decir que el consumo promedio de la población es menor a 25 dólares? Pruebe con α = 0.05.',
      partB: 'Calcule e interprete su intervalo de confianza al 97.5%.',
      method: 'z',
      tail: 'left',
      h0: 'μ ≥ 25',
      h1: 'μ < 25',
      alpha: 0.05,
      xbar: 22.6,
      mu0: 25,
      sigma: 2.5,
      sigmaLabel: 's',
      n: 49,
      intervals: [{ label: 'b', confidence: 0.975 }],
      finalTest: 'Sí existe evidencia de que el consumo promedio es menor a $25.',
      finalInterval: (interval) => `El consumo promedio se encuentra entre $${fmt(interval.li)} y $${fmt(interval.ls)} con 97.5% de confianza.`,
    },
    {
      number: 6,
      title: 'Ejercicio 6',
      statement:
        'Una muestra aleatoria de 12 alumnas graduadas de una escuela secretarial mecanografió un promedio de 79.5 palabras por minuto con una desviación estándar de 7.8 palabras por minuto.',
      partA: '¿Se tiene evidencia estadística para decir que el número promedio de palabras mecanografiadas por todas las graduadas de esa escuela es de 80 palabras por minuto con α = 0,01?',
      partB: 'Calcule e interprete su intervalo de confianza al 99%.',
      method: 't',
      tail: 'two',
      h0: 'μ = 80',
      h1: 'μ ≠ 80',
      alpha: 0.01,
      xbar: 79.5,
      mu0: 80,
      sigma: 7.8,
      sigmaLabel: 's',
      n: 12,
      intervals: [{ label: 'b', confidence: 0.99 }],
      finalTest: 'No se rechaza que el promedio sea de 80 palabras por minuto.',
      finalInterval: (interval) => `El promedio real se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} palabras por minuto con 99% de confianza.`,
    },
    {
      number: 7,
      title: 'Ejercicio 7',
      statement:
        'Un empresario está considerando la posibilidad de ampliar su negocio mediante la adquisición de un pequeño bar. El dueño actual del bar afirma que el ingreso diario del establecimiento sigue una distribución normal de media 675 soles y una desviación estándar de 75 soles. Para comprobar si decía la verdad, tomó una muestra de treinta días y ésta reveló un ingreso diario promedio de 625 soles.',
      partA: '¿Hay evidencia de que el ingreso diario promedio sea menor del que afirma el presente dueño? Utilice un nivel de significación del 1%.',
      partB: 'Calcula e interpreta un intervalo de confianza al 95%.',
      method: 'z',
      tail: 'left',
      h0: 'μ ≥ 675',
      h1: 'μ < 675',
      alpha: 0.01,
      xbar: 625,
      mu0: 675,
      sigma: 75,
      sigmaLabel: 'σ',
      n: 30,
      intervals: [{ label: 'b', confidence: 0.95 }],
      finalTest: 'Sí hay evidencia de que el ingreso diario promedio es menor que 675 soles.',
      finalInterval: (interval) => `El ingreso diario promedio se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} soles con 95% de confianza.`,
    },
    {
      number: 8,
      title: 'Ejercicio 8',
      statement:
        'Una empresa eléctrica fabrica focos que tienen una duración que está distribuida aproximadamente en forma normal con media de 800 horas y una desviación estándar de 40.',
      partA: 'Pruebe la hipótesis de que la duración promedio sea verdaderamente 800 horas si en una muestra aleatoria de 300 focos tiene una duración promedio de 788 horas. Utilice un nivel de significancia de 0.04.',
      partB: 'Calcula e interpreta un intervalo de confianza al 99%.',
      method: 'z',
      tail: 'two',
      h0: 'μ = 800',
      h1: 'μ ≠ 800',
      alpha: 0.04,
      xbar: 788,
      mu0: 800,
      sigma: 40,
      sigmaLabel: 'σ',
      n: 300,
      intervals: [{ label: 'b', confidence: 0.99 }],
      finalTest: 'Se rechaza que la duración promedio sea verdaderamente 800 horas.',
      finalInterval: (interval) => `La duración promedio se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} horas con 99% de confianza.`,
    },
    {
      number: 9,
      title: 'Ejercicio 9',
      statement:
        'Pruebe la hipótesis de que el contenido promedio en recipientes de un lubricante en particular es de 10 litros si los contenidos de una muestra aleatoria de 10 recipientes son 10.2, 9.7, 10.1, 10.3, 10.1, 9.8, 9.9, 10.4, 10.3 y 9.8 litros. Utilice un nivel de significancia de 0.01. Calcula e interpreta un intervalo de confianza al 95%.',
      partA: 'Pruebe la hipótesis de que el contenido promedio es de 10 litros con α = 0.01.',
      partB: 'Calcule e interprete un intervalo de confianza al 95%.',
      method: 't',
      tail: 'two',
      h0: 'μ = 10',
      h1: 'μ ≠ 10',
      alpha: 0.01,
      mu0: 10,
      values: [10.2, 9.7, 10.1, 10.3, 10.1, 9.8, 9.9, 10.4, 10.3, 9.8],
      sigmaLabel: 's',
      intervals: [{ label: 'b', confidence: 0.95 }],
      finalTest: 'No se rechaza que el contenido promedio sea de 10 litros.',
      finalInterval: (interval) => `El contenido promedio se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} litros con 95% de confianza.`,
    },
    {
      number: 10,
      title: 'Ejercicio 10',
      statement:
        'Una agencia de publicidad tiene un registro de datos sobre el tiempo (en minutos) de los anuncios publicitarios por cada 20 minutos en los programas principales de TV. Una muestra aleatoria de 35 de estos registros proporcionó un tiempo medio de publicidad de 3 minutos por cada 20 minutos de publicidad. Suponiendo que el tiempo de anuncios en minutos sigue una distribución normal con una desviación estándar de 1.2 minutos.',
      partA: 'Compruebe que el tiempo medio de publicidad sea mayor a los 2.5 minutos. Con una significancia del 3%.',
      partB: 'b) IC 99%.  c) IC 90%.  d) Tamaño de muestra con 95% de confianza y margen de error de 0.5 minutos.',
      method: 'z',
      tail: 'right',
      h0: 'μ ≤ 2.5',
      h1: 'μ > 2.5',
      alpha: 0.03,
      xbar: 3,
      mu0: 2.5,
      sigma: 1.2,
      sigmaLabel: 'σ',
      n: 35,
      intervals: [
        { label: 'b', confidence: 0.99 },
        { label: 'c', confidence: 0.9 },
      ],
      sampleSize: { confidence: 0.95, error: 0.5 },
      finalTest: 'Sí se comprueba que el tiempo medio de publicidad es mayor a 2.5 minutos.',
      finalInterval: (interval) => `IC ${fmt(interval.confidence * 100, interval.confidence === 0.9 ? 0 : 0)}%: ${fmt(interval.li)} a ${fmt(interval.ls)} minutos.`,
    },
    {
      number: 11,
      title: 'Ejercicio 11',
      statement:
        'El gerente de una distribuidora indica que el número promedio de llamadas por semana a los clientes es de 23 llamadas. Se pidió al personal de ventas que presentara informes semanales con los clientes llamados durante una semana. En una muestra de 36 informes semanales se determinó un promedio de 22.4 llamadas a clientes por semana, y una desviación estándar de 5 llamadas.',
      partA: 'Compruebe la afirmación del gerente con una significancia del 5%.',
      partB: 'Determinar un intervalo de confianza del 98% para el número promedio de llamadas semanales a clientes.',
      method: 'z',
      tail: 'two',
      h0: 'μ = 23',
      h1: 'μ ≠ 23',
      alpha: 0.05,
      xbar: 22.4,
      mu0: 23,
      sigma: 5,
      sigmaLabel: 's',
      n: 36,
      intervals: [{ label: 'b', confidence: 0.98 }],
      finalTest: 'No se rechaza la afirmación del gerente de 23 llamadas semanales.',
      finalInterval: (interval) => `El promedio de llamadas semanales se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} con 98% de confianza.`,
    },
    {
      number: 12,
      title: 'Ejercicio 12',
      statement:
        'Una empresa fabrica focos cuya duración tiene una distribución aproximadamente normal con desviación estándar poblacional de 60 horas. Suponga que una muestra de 20 focos tiene una duración promedio de 780 horas.',
      partA: 'Demuestre que la duración mínima promedio de los focos es de por lo menos 750 horas, con una significancia del 3%.',
      partB: 'Calcule e interprete un intervalo de confianza del 96% para la duración promedio de todos los focos producidos por esta empresa.',
      method: 'z',
      tail: 'right',
      h0: 'μ ≤ 750',
      h1: 'μ > 750',
      alpha: 0.03,
      xbar: 780,
      mu0: 750,
      sigma: 60,
      sigmaLabel: 'σ',
      n: 20,
      intervals: [{ label: 'b', confidence: 0.96 }],
      finalTest: 'Sí hay evidencia para afirmar que la duración promedio es por lo menos 750 horas.',
      finalInterval: (interval) => `La duración promedio se encuentra entre ${fmt(interval.li)} y ${fmt(interval.ls)} horas con 96% de confianza.`,
    },
  ];
  return base.map(calcExercise);
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

function box(ws, row1, col1, row2, col2) {
  for (let row = row1; row <= row2; row += 1) {
    for (let col = col1; col <= col2; col += 1) {
      ws.getCell(row, col).border = border;
    }
  }
}

function setupSheet(ws) {
  ws.views = [{ showGridLines: true }];
  const widths = [3, 5, 15, 12, 14, 3, 12, 13, 14, 3, 16, 11, 15, 22];
  widths.forEach((width, index) => {
    ws.getColumn(index + 1).width = width;
  });
  ws.properties.defaultRowHeight = 19;
}

function valueFormula(value, formula, numFmt = '0.000000') {
  return { formula, result: value };
}

function criticalFormula(ex, alphaCell) {
  if (ex.method === 'z') {
    if (ex.tail === 'right') return `_xlfn.NORM.S.INV(1-${alphaCell})`;
    if (ex.tail === 'left') return `_xlfn.NORM.S.INV(${alphaCell})`;
    return `_xlfn.NORM.S.INV(1-${alphaCell}/2)`;
  }
  return `TINV(${alphaCell},${ex.df})`;
}

function pValueFormula(ex, statCell) {
  if (ex.method === 'z') {
    if (ex.tail === 'right') return `1-_xlfn.NORM.S.DIST(${statCell},1)`;
    if (ex.tail === 'left') return `_xlfn.NORM.S.DIST(${statCell},1)`;
    return `2*(1-_xlfn.NORM.S.DIST(ABS(${statCell}),1))`;
  }
  return `TDIST(ABS(${statCell}),${ex.df},2)`;
}

function addTop(ws, ex) {
  merge(ws, 2, 2, 2, 14, ex.title, {
    bold: true,
    size: 22,
    color: colors.red,
    align: 'left',
    border: false,
  });
  merge(ws, 4, 2, 7, 14, ex.statement, {
    fill: colors.peach,
    wrap: true,
    border: true,
  });
  ws.getRow(4).height = 28;
  ws.getRow(5).height = 28;
  ws.getRow(6).height = 28;
  ws.getRow(7).height = 28;
  setCell(ws, 9, 2, 'a)', { bold: true, color: '1F4E79', border: false });
  merge(ws, 9, 3, 10, 14, ex.partA, { border: false });
  setCell(ws, 11, 2, ex.number === 10 ? 'b-c-d)' : 'b)', { bold: true, color: '1F4E79', border: false });
  merge(ws, 11, 3, 12, 14, ex.partB, { border: false });
}

function addTestBlock(ws, ex) {
  merge(ws, 15, 2, 15, 7, 'A) Prueba de hipótesis', { bold: true, align: 'center', fill: colors.blue });
  merge(ws, 16, 2, 16, 4, 'Elemento', { bold: true, align: 'center', fill: colors.beige });
  setCell(ws, 16, 5, 'Símbolo', { bold: true, align: 'center', fill: colors.beige });
  merge(ws, 16, 6, 16, 7, 'Valor', { bold: true, align: 'center', fill: colors.beige });

  const rows = [
    ['Hipótesis nula', 'H0', ex.h0, ''],
    ['Hipótesis alternativa', 'H1', ex.h1, ''],
    ['Nivel de significancia', 'α', ex.alpha, '0.00'],
    ['Media muestral', 'x̄', ex.values ? { formula: 'AVERAGE(B36:B45)', result: ex.xbar } : ex.xbar, '0.00'],
    ['Media hipotética', 'µ0', ex.mu0, '0.00'],
    ['Desv. estándar', ex.sigmaLabel, ex.values ? { formula: 'STDEV(B36:B45)', result: ex.sigma } : ex.sigma, '0.00'],
    ['Tamaño de muestra', 'n', ex.n, '0'],
  ];

  if (ex.method === 't') {
    rows.push(['Grados de libertad', 'gl', ex.df, '0']);
  }

  rows.forEach((item, index) => {
    const row = 17 + index;
    merge(ws, row, 2, row, 4, item[0], { fill: colors.gray });
    setCell(ws, row, 5, item[1], { align: 'center' });
    merge(ws, row, 6, row, 7, item[2], { align: 'right', numFmt: item[3] || undefined });
  });

  const baseRow = 17;
  const alphaRow = baseRow + 2;
  const xbarRow = baseRow + 3;
  const mu0Row = baseRow + 4;
  const sigmaRow = baseRow + 5;
  const nRow = baseRow + 6;
  const statRow = ex.method === 't' ? 26 : 25;
  const criticalRow = statRow + 1;
  const pRow = statRow + 2;
  const decisionRow = statRow + 3;
  const responseRow = statRow + 5;

  const xbarCell = ws.getCell(xbarRow, 6).address;
  const mu0Cell = ws.getCell(mu0Row, 6).address;
  const sigmaCell = ws.getCell(sigmaRow, 6).address;
  const nCell = ws.getCell(nRow, 6).address;
  const alphaCell = ws.getCell(alphaRow, 6).address;

  merge(ws, statRow, 2, statRow, 4, `Estadístico ${ex.method.toUpperCase()} calculado`, { fill: colors.gray });
  setCell(ws, statRow, 5, `${ex.method.toUpperCase()}c`, { align: 'center' });
  merge(ws, statRow, 6, statRow, 7, valueFormula(ex.stat, `(${xbarCell}-${mu0Cell})/(${sigmaCell}/SQRT(${nCell}))`), {
    align: 'right',
    numFmt: '0.000000',
  });

  const statCell = ws.getCell(statRow, 6).address;
  merge(ws, criticalRow, 2, criticalRow, 4, `Valor crítico ${ex.method.toUpperCase()}`, { fill: colors.gray });
  setCell(ws, criticalRow, 5, ex.tail === 'two' ? `±${ex.method.toUpperCase()}t` : `${ex.method.toUpperCase()}t`, { align: 'center' });
  merge(ws, criticalRow, 6, criticalRow, 7, valueFormula(ex.critical, criticalFormula(ex, alphaCell)), {
    align: 'right',
    numFmt: '0.000000',
  });

  merge(ws, pRow, 2, pRow, 4, 'p-valor', { fill: colors.gray });
  setCell(ws, pRow, 5, 'p', { align: 'center' });
  merge(ws, pRow, 6, pRow, 7, valueFormula(ex.pvalue, pValueFormula(ex, statCell)), {
    align: 'right',
    numFmt: '0.000000',
  });

  merge(ws, decisionRow, 2, decisionRow, 4, 'Decisión', { fill: colors.beige, bold: true });
  merge(ws, decisionRow, 5, decisionRow, 7, ex.decision, {
    fill: colors.green,
    bold: true,
    align: 'center',
    color: colors.red,
  });

  merge(ws, responseRow, 2, responseRow, 4, 'Respuesta final', { fill: colors.beige, bold: true });
  merge(ws, responseRow, 5, responseRow + 2, 7, ex.finalTest, {
    fill: colors.green,
    bold: true,
    align: 'center',
    color: colors.red,
  });

  if (ex.values) {
    merge(ws, 35, 2, 35, 7, 'Datos originales del ejercicio', { fill: colors.beige, bold: true });
    ex.values.forEach((value, index) => {
      setCell(ws, 36 + index, 2, value, { align: 'center', numFmt: '0.0' });
      merge(ws, 36 + index, 3, 36 + index, 7, `Dato ${index + 1}`, { fill: colors.gray });
    });
    box(ws, 35, 2, 45, 7);
  }

  box(ws, 15, 2, responseRow + 2, 7);
}

function addIntervalBlock(ws, ex, interval, startRow, title) {
  merge(ws, startRow, 9, startRow, 14, title, { bold: true, align: 'center', fill: colors.blue });
  merge(ws, startRow + 1, 9, startRow + 1, 11, 'Elemento', { bold: true, align: 'center', fill: colors.beige });
  setCell(ws, startRow + 1, 12, 'Símbolo', { bold: true, align: 'center', fill: colors.beige });
  merge(ws, startRow + 1, 13, startRow + 1, 14, 'Valor', { bold: true, align: 'center', fill: colors.beige });

  const rows = [
    ['Nivel de confianza', 'NC', interval.confidence, '0.0%'],
    ['Media muestral', 'x̄', ex.xbar, '0.00'],
    ['Desv. estándar', ex.sigmaLabel, ex.sigma, '0.00'],
    ['Tamaño de muestra', 'n', ex.n, '0'],
  ];

  if (ex.method === 't') rows.push(['Grados de libertad', 'gl', ex.df, '0']);

  rows.forEach((item, index) => {
    const row = startRow + 2 + index;
    merge(ws, row, 9, row, 11, item[0], { fill: colors.gray });
    setCell(ws, row, 12, item[1], { align: 'center' });
    merge(ws, row, 13, row, 14, item[2], { align: 'right', numFmt: item[3] });
  });

  const confidenceCell = ws.getCell(startRow + 2, 13).address;
  const xbarCell = ws.getCell(startRow + 3, 13).address;
  const sigmaCell = ws.getCell(startRow + 4, 13).address;
  const nCell = ws.getCell(startRow + 5, 13).address;
  const criticalRow = ex.method === 't' ? startRow + 7 : startRow + 6;

  merge(ws, criticalRow, 9, criticalRow, 11, `Valor ${ex.method.toUpperCase()}`, { fill: colors.gray });
  setCell(ws, criticalRow, 12, ex.method.toUpperCase(), { align: 'center' });
  const critFormula =
    ex.method === 'z'
      ? `_xlfn.NORM.S.INV((1+${confidenceCell})/2)`
      : `TINV(1-${confidenceCell},${ex.df})`;
  merge(ws, criticalRow, 13, criticalRow, 14, valueFormula(interval.critical, critFormula), {
    align: 'right',
    numFmt: '0.000000',
  });

  const criticalCell = ws.getCell(criticalRow, 13).address;
  const errorRow = criticalRow + 1;
  merge(ws, errorRow, 9, errorRow, 11, 'Error', { fill: colors.gray });
  setCell(ws, errorRow, 12, 'E', { align: 'center' });
  merge(ws, errorRow, 13, errorRow, 14, valueFormula(interval.error, `${criticalCell}*${sigmaCell}/SQRT(${nCell})`), {
    align: 'right',
    numFmt: '0.000000',
  });

  const errorCell = ws.getCell(errorRow, 13).address;
  const liRow = errorRow + 1;
  merge(ws, liRow, 9, liRow, 11, 'Límite inferior', { fill: colors.gray });
  setCell(ws, liRow, 12, 'LI', { align: 'center' });
  merge(ws, liRow, 13, liRow, 14, valueFormula(interval.li, `${xbarCell}-${errorCell}`), {
    align: 'right',
    numFmt: '0.000000',
  });

  const lsRow = liRow + 1;
  merge(ws, lsRow, 9, lsRow, 11, 'Límite superior', { fill: colors.gray });
  setCell(ws, lsRow, 12, 'LS', { align: 'center' });
  merge(ws, lsRow, 13, lsRow, 14, valueFormula(interval.ls, `${xbarCell}+${errorCell}`), {
    align: 'right',
    numFmt: '0.000000',
  });

  const responseRow = lsRow + 2;
  merge(ws, responseRow, 9, responseRow, 11, 'Respuesta final', { fill: colors.beige, bold: true });
  merge(ws, responseRow, 12, responseRow + 2, 14, ex.finalInterval(interval), {
    fill: colors.green,
    bold: true,
    align: 'center',
    color: colors.red,
  });

  box(ws, startRow, 9, responseRow + 2, 14);
  return responseRow + 3;
}

function addSampleSizeBlock(ws, ex, startRow) {
  const spec = ex.sampleSize;
  if (!spec) return;
  const z = normInv((1 + spec.confidence) / 2);
  const rawN = (z * ex.sigma / spec.error) ** 2;
  const finalN = Math.ceil(rawN);

  merge(ws, startRow, 9, startRow, 14, 'Ejercicio 10.d: Tamaño de muestra', {
    bold: true,
    align: 'center',
    fill: colors.blue,
  });
  const rows = [
    ['Nivel de confianza', 'NC', spec.confidence, '0%'],
    ['Valor Z', 'Z', { formula: `_xlfn.NORM.S.INV((1+M${startRow + 1})/2)`, result: z }, '0.000000'],
    ['Desv. estándar', 'σ', ex.sigma, '0.00'],
    ['Margen de error', 'E', spec.error, '0.00'],
    ['Tamaño calculado', 'n', { formula: `(M${startRow + 2}*M${startRow + 3}/M${startRow + 4})^2`, result: rawN }, '0.000000'],
    ['Tamaño mínimo', 'n', finalN, '0'],
  ];

  rows.forEach((item, index) => {
    const row = startRow + 1 + index;
    merge(ws, row, 9, row, 11, item[0], { fill: colors.gray });
    setCell(ws, row, 12, item[1], { align: 'center' });
    merge(ws, row, 13, row, 14, item[2], { align: 'right', numFmt: item[3] });
  });

  merge(ws, startRow + 8, 9, startRow + 8, 11, 'Respuesta final', { fill: colors.beige, bold: true });
  merge(ws, startRow + 8, 12, startRow + 10, 14, `Se debe tomar una muestra mínima de ${finalN} registros.`, {
    fill: colors.green,
    bold: true,
    align: 'center',
    color: colors.red,
  });
  box(ws, startRow, 9, startRow + 10, 14);
}

function addExerciseSheet(workbook, ex) {
  const ws = workbook.addWorksheet(`Ejercicio ${ex.number}`);
  setupSheet(ws);
  addTop(ws, ex);
  addTestBlock(ws, ex);
  let nextRow = 15;
  ex.intervals.forEach((interval, index) => {
    const title = ex.number === 10 ? `Ejercicio 10.${interval.label}: Intervalo de confianza` : 'B) Intervalo de confianza';
    nextRow = addIntervalBlock(ws, ex, interval, nextRow, title) + 1;
  });
  addSampleSizeBlock(ws, ex, nextRow);
  ws.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { left: 0.25, right: 0.25, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 },
  };
}

function addSummarySheet(workbook, data) {
  const ws = workbook.addWorksheet('Resumen');
  setupSheet(ws);
  merge(ws, 2, 2, 2, 14, 'Estimación y prueba de hipótesis de una media', {
    bold: true,
    size: 14,
    align: 'center',
    border: false,
  });
  merge(ws, 3, 2, 3, 14, 'Resumen rápido de decisiones e intervalos de confianza.', {
    italic: true,
    align: 'center',
    border: false,
  });
  merge(ws, 5, 2, 5, 14, 'Resumen rápido', { bold: true, align: 'center', fill: colors.beige });
  setCell(ws, 6, 2, 'Ejercicio', { bold: true, align: 'center', fill: colors.blue });
  merge(ws, 6, 3, 6, 4, 'Estadístico', { bold: true, align: 'center', fill: colors.blue });
  merge(ws, 6, 5, 6, 7, 'Decisión', { bold: true, align: 'center', fill: colors.blue });
  merge(ws, 6, 8, 6, 14, 'Resultado / intervalo', { bold: true, align: 'center', fill: colors.blue });

  data.forEach((ex, index) => {
    const row = 7 + index;
    setCell(ws, row, 2, `Ejercicio ${ex.number}`, { align: 'center', fill: colors.gray });
    merge(ws, row, 3, row, 4, fmt(ex.stat, 4), { align: 'center' });
    merge(ws, row, 5, row, 7, ex.decision, { align: 'center', fill: colors.green, bold: true, color: colors.red });
    const intervalText = ex.intervals
      .map((interval) => `IC ${fmt(interval.confidence * 100, interval.confidence === 0.975 ? 1 : 0)}%: ${fmt(interval.li)} a ${fmt(interval.ls)}`)
      .join(' | ');
    merge(ws, row, 8, row, 14, `${ex.finalTest} ${intervalText}`, { fill: colors.green, color: colors.red, bold: true });
  });
}

async function main() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Codex';
  workbook.created = new Date();
  const data = makeData();
  addSummarySheet(workbook, data);
  data.forEach((ex) => addExerciseSheet(workbook, ex));
  await workbook.xlsx.writeFile(outputPath);
  console.log(outputPath);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
