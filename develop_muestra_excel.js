const ExcelJS = require('exceljs');

const INPUT_PATH = 'C:/Users/USER/Downloads/Actividad de seleccion de muestra (2).xlsx';
const DEFAULT_OUTPUT_PATH = 'D:/Estadistica Inferencial/Actividad de seleccion de muestra (2)_desarrollado.xlsx';

const COLORS = {
  black: 'FF000000',
  white: 'FFFFFFFF',
  red: 'FFFF0000',
  darkRed: 'FFC00000',
  grayFill: 'FFF2F2F2',
  greenFill: 'FFE2F0D9',
  blueFill: 'FFDDEBF7',
  yellowFill: 'FFFFF2CC',
  orangeFill: 'FFF4B183',
};

function mulberry32(seed) {
  let value = seed >>> 0;
  return () => {
    value += 0x6d2b79f5;
    let temp = Math.imul(value ^ (value >>> 15), value | 1);
    temp ^= temp + Math.imul(temp ^ (temp >>> 7), temp | 61);
    return ((temp ^ (temp >>> 14)) >>> 0) / 4294967296;
  };
}

function randomInt(rng, min, max) {
  return Math.floor(rng() * (max - min + 1)) + min;
}

function setThinBorder(cell) {
  cell.border = {
    top: { style: 'thin', color: { argb: COLORS.black } },
    left: { style: 'thin', color: { argb: COLORS.black } },
    bottom: { style: 'thin', color: { argb: COLORS.black } },
    right: { style: 'thin', color: { argb: COLORS.black } },
  };
}

function borderRange(worksheet, startRow, endRow, startCol, endCol) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      setThinBorder(worksheet.getCell(row, col));
    }
  }
}

function fillRange(worksheet, startRow, endRow, startCol, endCol, color) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      worksheet.getCell(row, col).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: color },
      };
    }
  }
}

function centerRange(worksheet, startRow, endRow, startCol, endCol) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      worksheet.getCell(row, col).alignment = {
        vertical: 'middle',
        horizontal: 'center',
      };
    }
  }
}

function applyHeaderStyle(cell, color = COLORS.grayFill) {
  cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.black } };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
  cell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: color },
  };
  setThinBorder(cell);
}

function asText(value) {
  if (value == null) {
    return '';
  }
  if (typeof value === 'object' && value.text) {
    return value.text;
  }
  return String(value);
}

function setFormula(cell, formula, result, numFmt) {
  cell.value = { formula, result };
  if (numFmt) {
    cell.numFmt = numFmt;
  }
}

function clearWorksheet(worksheet) {
  worksheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.value = null;
      cell.style = {};
    });
  });

  while (worksheet._merges && worksheet._merges.size > 0) {
    const merge = worksheet._merges.keys().next().value;
    worksheet.unMergeCells(merge);
  }
}

function setStandardFont(worksheet, startRow, endRow, startCol, endCol) {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      worksheet.getCell(row, col).font = worksheet.getCell(row, col).font || {
        name: 'Calibri',
        size: 11,
      };
    }
  }
}

function buildActividadSheet(worksheet, stats) {
  worksheet.columns = [
    { key: 'a', width: 4 },
    { key: 'b', width: 30 },
    { key: 'c', width: 15 },
    { key: 'd', width: 15 },
    { key: 'e', width: 12 },
    { key: 'f', width: 12 },
    { key: 'g', width: 8 },
    { key: 'h', width: 8 },
    { key: 'i', width: 30 },
    { key: 'j', width: 15 },
    { key: 'k', width: 15 },
    { key: 'l', width: 12 },
    { key: 'm', width: 12 },
    { key: 'n', width: 8 },
    { key: 'o', width: 12 },
    { key: 'p', width: 12 },
  ];

  worksheet.getCell('B2').font = { name: 'Calibri', size: 18, bold: true };
  worksheet.getCell('B2').alignment = { vertical: 'middle', horizontal: 'left' };

  worksheet.getCell('B11').value = 'Pregunta 1';
  worksheet.getCell('B11').font = { name: 'Calibri', size: 12, bold: true, italic: true, color: { argb: COLORS.red } };

  worksheet.getCell('I11').value = 'Pregunta 2';
  worksheet.getCell('I11').font = { name: 'Calibri', size: 12, bold: true, italic: true, color: { argb: COLORS.red } };

  worksheet.mergeCells('B13:D13');
  worksheet.getCell('B13').value = 'Tamaño de la muestra';
  worksheet.getCell('B13').font = { name: 'Calibri', size: 11, bold: true };

  worksheet.mergeCells('E13:G13');
  worksheet.getCell('E13').value = {
    richText: [
      { text: 'Variable: ', font: { name: 'Calibri', size: 11 } },
      { text: 'Impuesto', font: { name: 'Calibri', size: 11, color: { argb: COLORS.red } } },
    ],
  };

  worksheet.mergeCells('I13:K13');
  worksheet.getCell('I13').value = 'Tamaño de la muestra';
  worksheet.getCell('I13').font = { name: 'Calibri', size: 11, bold: true };

  worksheet.mergeCells('L13:P13');
  worksheet.getCell('L13').value = {
    richText: [
      { text: 'Variable: ', font: { name: 'Calibri', size: 11 } },
      {
        text: 'Empresas que están de acuerdo con el nuevo impuesto',
        font: { name: 'Calibri', size: 11, color: { argb: COLORS.red } },
      },
    ],
  };

  const leftLabels = [
    ['B15', 'Tamaño de la población'],
    ['B16', 'Nivel de confianza:'],
    ['B17', 'α ='],
    ['B18', 'Margen de Error:'],
    ['B19', 'Desviación estánd. Pob.'],
    ['B20', 'Tamaño de la muestra:'],
  ];

  const rightLabels = [
    ['I15', 'Tamaño de la población'],
    ['I16', 'Nivel de confianza:'],
    ['I17', 'Margen de error:'],
    ['I18', 'Proporción esperada'],
    ['I19', 'Complemento'],
    ['I20', 'Tamaño de la muestra:'],
  ];

  for (const [address, value] of leftLabels) {
    worksheet.getCell(address).value = value;
  }
  for (const [address, value] of rightLabels) {
    worksheet.getCell(address).value = value;
  }

  worksheet.getCell('E15').value = 'N =';
  worksheet.getCell('E16').value = '1 - α =';
  worksheet.getCell('E17').value = 'nivel de error';
  worksheet.getCell('E18').value = 'e =';
  worksheet.getCell('E19').value = 'σ =';
  worksheet.getCell('E20').value = 'n =';

  setFormula(worksheet.getCell('F15'), 'COUNTA(MATRIZ!A:A)-1', stats.N, '0');
  worksheet.getCell('F16').value = stats.conf1;
  worksheet.getCell('F16').numFmt = '0.00';
  setFormula(worksheet.getCell('F17'), '1-F16', stats.alpha1, '0.00');
  worksheet.getCell('F18').value = stats.error1;
  worksheet.getCell('F18').numFmt = '0';
  setFormula(worksheet.getCell('F19'), 'SQRT(122500)', stats.sigma, '0');
  setFormula(worksheet.getCell('F20'), '(F15*(I16^2)*(F19^2))/(((F15-1)*(F18^2))+((I16^2)*(F19^2)))', stats.n1Raw, '0.000000');
  setFormula(worksheet.getCell('G20'), 'ROUNDUP(F20,0)', stats.n1, '0');
  worksheet.getCell('G20').font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };

  worksheet.getCell('H16').value = 'Z =';
  setFormula(worksheet.getCell('I16'), 'NORMSINV((1+F16)/2)', stats.z1, '0.000');
  worksheet.getCell('J16').value = '=INV.NORM.ESTAND(0.96)';
  worksheet.getCell('J16').font = { name: 'Calibri', size: 10, italic: true };

  worksheet.getCell('L15').value = 'N =';
  worksheet.getCell('L16').value = '1 - α =';
  worksheet.getCell('L17').value = 'e =';
  worksheet.getCell('L18').value = 'P =';
  worksheet.getCell('L19').value = 'q =';
  worksheet.getCell('L20').value = 'n =';

  setFormula(worksheet.getCell('M15'), 'COUNTA(MATRIZ!A:A)-1', stats.N, '0');
  worksheet.getCell('M16').value = stats.conf2;
  worksheet.getCell('M16').numFmt = '0.0';
  worksheet.getCell('M17').value = stats.error2;
  worksheet.getCell('M17').numFmt = '0.00';
  worksheet.getCell('M18').value = stats.p;
  worksheet.getCell('M18').numFmt = '0.00';
  setFormula(worksheet.getCell('M19'), '1-M18', stats.q, '0.00');
  setFormula(worksheet.getCell('M20'), '(M15*(P16^2)*M18*M19)/(((M15-1)*(M17^2))+((P16^2)*M18*M19))', stats.n2Raw, '0.000000');
  setFormula(worksheet.getCell('N20'), 'ROUNDUP(M20,0)', stats.n2, '0');
  worksheet.getCell('N20').font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };

  worksheet.getCell('O16').value = 'Z =';
  setFormula(worksheet.getCell('P16'), 'NORMSINV((1+M16)/2)', stats.z2, '0.0000');

  worksheet.mergeCells('B22:G22');
  worksheet.getCell('B22').value = 'n = (N·Z²·σ²) / ((N - 1)·e² + Z²·σ²)';
  worksheet.getCell('B22').font = { name: 'Calibri', size: 10, italic: true };
  worksheet.getCell('B22').alignment = { horizontal: 'center' };

  worksheet.mergeCells('I22:P22');
  worksheet.getCell('I22').value = 'n = (N·Z²·p·q) / ((N - 1)·e² + Z²·p·q)';
  worksheet.getCell('I22').font = { name: 'Calibri', size: 10, italic: true };
  worksheet.getCell('I22').alignment = { horizontal: 'center' };

  worksheet.mergeCells('B25:M25');
  worksheet.getCell('B25').value = 'Para las preguntas 3 y 4 se utiliza el mayor tamaño de muestra calculado.';
  worksheet.getCell('B25').font = { name: 'Calibri', size: 11, bold: true };
  worksheet.getCell('B25').alignment = { horizontal: 'center' };
  worksheet.getCell('N25').value = 'n =';
  setFormula(worksheet.getCell('O25'), 'MAX(G20,N20)', stats.nFinal, '0');
  worksheet.getCell('O25').font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };

  borderRange(worksheet, 15, 20, 2, 7);
  borderRange(worksheet, 16, 16, 8, 9);
  borderRange(worksheet, 15, 20, 9, 14);
  borderRange(worksheet, 16, 16, 15, 16);
  borderRange(worksheet, 25, 25, 2, 15);

  fillRange(worksheet, 15, 20, 2, 2, COLORS.grayFill);
  fillRange(worksheet, 15, 20, 9, 9, COLORS.grayFill);
  fillRange(worksheet, 25, 25, 2, 15, COLORS.yellowFill);
  centerRange(worksheet, 15, 20, 5, 7);
  centerRange(worksheet, 15, 20, 12, 14);
  centerRange(worksheet, 16, 16, 8, 9);
  centerRange(worksheet, 16, 16, 15, 16);

  setStandardFont(worksheet, 11, 25, 2, 16);
  worksheet.getCell('B11').font = { name: 'Calibri', size: 12, bold: true, italic: true, color: { argb: COLORS.red } };
  worksheet.getCell('I11').font = { name: 'Calibri', size: 12, bold: true, italic: true, color: { argb: COLORS.red } };
}

function buildMasSheet(worksheet, data, stats) {
  clearWorksheet(worksheet);

  worksheet.columns = [
    { key: 'id', width: 8 },
    { key: 'razon', width: 42 },
    { key: 'sector', width: 10 },
    { key: 'aleatorio', width: 12 },
    { key: 'space', width: 3 },
    { key: 'f', width: 8 },
    { key: 'g', width: 34 },
    { key: 'h', width: 10 },
    { key: 'i', width: 12 },
    { key: 'j', width: 10 },
    { key: 'k', width: 10 },
    { key: 'l', width: 12 },
    { key: 'm', width: 12 },
  ];

  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  worksheet.getCell('A1').value = 'Id';
  worksheet.getCell('B1').value = 'RAZÓN SOCIAL';
  worksheet.getCell('C1').value = 'SECTOR';
  worksheet.getCell('D1').value = 'ALEATORIO';

  ['A1', 'B1', 'C1', 'D1'].forEach((address) => applyHeaderStyle(worksheet.getCell(address), COLORS.blueFill));

  data.masSorted.forEach((row, index) => {
    const targetRow = index + 2;
    worksheet.getCell(`A${targetRow}`).value = row.id;
    worksheet.getCell(`B${targetRow}`).value = row.name;
    worksheet.getCell(`C${targetRow}`).value = row.sector;
    worksheet.getCell(`D${targetRow}`).value = row.random;
    worksheet.getCell(`A${targetRow}`).numFmt = '0';
    worksheet.getCell(`C${targetRow}`).numFmt = '0';
    worksheet.getCell(`D${targetRow}`).numFmt = '0';
  });

  worksheet.autoFilter = 'A1:D1197';

  worksheet.mergeCells('F1:M1');
  worksheet.getCell('F1').value = 'Selección de muestra con muestreo aleatorio simple';
  worksheet.getCell('F1').font = { name: 'Calibri', size: 13, bold: true };
  worksheet.getCell('F1').alignment = { horizontal: 'center' };

  worksheet.mergeCells('F2:H2');
  worksheet.getCell('F2').value = 'Tamaño de la muestra';
  worksheet.getCell('F2').font = { name: 'Calibri', size: 11, bold: true };

  worksheet.mergeCells('I2:M2');
  worksheet.getCell('I2').value = {
    richText: [
      { text: 'Variable guía: ', font: { name: 'Calibri', size: 11 } },
      {
        text: 'Empresas que están de acuerdo con el nuevo impuesto',
        font: { name: 'Calibri', size: 11, color: { argb: COLORS.red } },
      },
    ],
  };

  worksheet.mergeCells('F3:M3');
  worksheet.getCell('F3').value = 'Se utiliza el mayor tamaño calculado entre la pregunta 1 y la pregunta 2.';
  worksheet.getCell('F3').font = { name: 'Calibri', size: 10, italic: true };

  const labels = [
    ['F5', 'Tamaño de la población'],
    ['F6', 'Nivel de confianza:'],
    ['F7', 'Margen de error:'],
    ['F8', 'P ='],
    ['F9', 'q ='],
    ['F10', 'Tamaño de la muestra:'],
  ];

  for (const [address, label] of labels) {
    worksheet.getCell(address).value = label;
  }

  worksheet.getCell('I5').value = 'N =';
  worksheet.getCell('I6').value = '1 - α =';
  worksheet.getCell('I7').value = 'e =';
  worksheet.getCell('I8').value = 'P =';
  worksheet.getCell('I9').value = 'q =';
  worksheet.getCell('I10').value = 'n =';

  setFormula(worksheet.getCell('J5'), 'Actividad!M15', stats.N, '0');
  setFormula(worksheet.getCell('J6'), 'Actividad!M16', stats.conf2, '0.0');
  setFormula(worksheet.getCell('J7'), 'Actividad!M17', stats.error2, '0.00');
  setFormula(worksheet.getCell('J8'), 'Actividad!M18', stats.p, '0.00');
  setFormula(worksheet.getCell('J9'), 'Actividad!M19', stats.q, '0.00');
  setFormula(worksheet.getCell('J10'), 'Actividad!M20', stats.n2Raw, '0.000000');
  setFormula(worksheet.getCell('K10'), 'Actividad!N20', stats.n2, '0');
  worksheet.getCell('K10').font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };

  worksheet.getCell('L6').value = 'Z =';
  setFormula(worksheet.getCell('M6'), 'Actividad!P16', stats.z2, '0.0000');

  borderRange(worksheet, 5, 10, 6, 11);
  borderRange(worksheet, 6, 6, 12, 13);
  fillRange(worksheet, 5, 10, 6, 6, COLORS.grayFill);
  centerRange(worksheet, 5, 10, 9, 11);
  centerRange(worksheet, 6, 6, 12, 13);

  worksheet.getCell('F12').value = 'Selección de la muestra';
  worksheet.getCell('F12').font = { name: 'Calibri', size: 11, bold: true, italic: true };

  ['F13', 'G13', 'H13', 'I13'].forEach((address) => applyHeaderStyle(worksheet.getCell(address), COLORS.greenFill));
  worksheet.getCell('F13').value = 'N°';
  worksheet.getCell('G13').value = 'RAZÓN SOCIAL';
  worksheet.getCell('H13').value = 'SECTOR';
  worksheet.getCell('I13').value = 'ALEATORIO';

  data.masSample.forEach((row, index) => {
    const targetRow = index + 14;
    worksheet.getCell(`F${targetRow}`).value = index + 1;
    setFormula(worksheet.getCell(`G${targetRow}`), `B${index + 2}`, row.name);
    setFormula(worksheet.getCell(`H${targetRow}`), `C${index + 2}`, row.sector);
    setFormula(worksheet.getCell(`I${targetRow}`), `D${index + 2}`, row.random);
    worksheet.getCell(`F${targetRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`H${targetRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`I${targetRow}`).alignment = { horizontal: 'center' };
  });

  borderRange(worksheet, 13, 13 + data.masSample.length, 6, 9);
}

function buildMsSheet(worksheet, data, stats) {
  clearWorksheet(worksheet);

  worksheet.columns = [
    { key: 'id', width: 8 },
    { key: 'razon', width: 40 },
    { key: 'sector', width: 10 },
    { key: 'space', width: 3 },
    { key: 'e', width: 8 },
    { key: 'f', width: 8 },
    { key: 'g', width: 34 },
    { key: 'h', width: 10 },
    { key: 'i', width: 12 },
    { key: 'j', width: 10 },
    { key: 'k', width: 12 },
  ];

  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  worksheet.getCell('A1').value = 'Id';
  worksheet.getCell('B1').value = 'RAZÓN SOCIAL';
  worksheet.getCell('C1').value = 'SECTOR';
  ['A1', 'B1', 'C1'].forEach((address) => applyHeaderStyle(worksheet.getCell(address), COLORS.blueFill));

  data.population.forEach((row, index) => {
    const targetRow = index + 2;
    worksheet.getCell(`A${targetRow}`).value = row.id;
    worksheet.getCell(`B${targetRow}`).value = row.name;
    worksheet.getCell(`C${targetRow}`).value = row.sector;
  });

  worksheet.autoFilter = 'A1:C1197';

  worksheet.mergeCells('E1:K1');
  worksheet.getCell('E1').value = 'Selección de muestra con muestreo sistemático';
  worksheet.getCell('E1').font = { name: 'Calibri', size: 13, bold: true };
  worksheet.getCell('E1').alignment = { horizontal: 'center' };

  worksheet.mergeCells('E2:G2');
  worksheet.getCell('E2').value = 'Tamaño de la muestra';
  worksheet.getCell('E2').font = { name: 'Calibri', size: 11, bold: true };

  worksheet.mergeCells('H2:K2');
  worksheet.getCell('H2').value = {
    richText: [
      { text: 'Variable guía: ', font: { name: 'Calibri', size: 11 } },
      {
        text: 'Empresas que están de acuerdo con el nuevo impuesto',
        font: { name: 'Calibri', size: 11, color: { argb: COLORS.red } },
      },
    ],
  };

  const labels = [
    ['E4', 'Tamaño de la población'],
    ['E5', 'Nivel de confianza:'],
    ['E6', 'Margen de error:'],
    ['E7', 'P ='],
    ['E8', 'q ='],
    ['E9', 'Tamaño de la muestra:'],
  ];

  for (const [address, label] of labels) {
    worksheet.getCell(address).value = label;
  }

  worksheet.getCell('H4').value = 'N =';
  worksheet.getCell('H5').value = '1 - α =';
  worksheet.getCell('H6').value = 'e =';
  worksheet.getCell('H7').value = 'P =';
  worksheet.getCell('H8').value = 'q =';
  worksheet.getCell('H9').value = 'n =';

  setFormula(worksheet.getCell('I4'), 'Actividad!M15', stats.N, '0');
  setFormula(worksheet.getCell('I5'), 'Actividad!M16', stats.conf2, '0.0');
  setFormula(worksheet.getCell('I6'), 'Actividad!M17', stats.error2, '0.00');
  setFormula(worksheet.getCell('I7'), 'Actividad!M18', stats.p, '0.00');
  setFormula(worksheet.getCell('I8'), 'Actividad!M19', stats.q, '0.00');
  setFormula(worksheet.getCell('I9'), 'Actividad!M20', stats.n2Raw, '0.000000');
  setFormula(worksheet.getCell('J9'), 'Actividad!N20', stats.n2, '0');
  worksheet.getCell('J9').font = { name: 'Calibri', size: 11, bold: true, color: { argb: COLORS.red } };

  worksheet.getCell('J5').value = 'Z =';
  setFormula(worksheet.getCell('K5'), 'Actividad!P16', stats.z2, '0.0000');

  borderRange(worksheet, 4, 9, 5, 10);
  borderRange(worksheet, 5, 5, 10, 11);
  fillRange(worksheet, 4, 9, 5, 5, COLORS.grayFill);
  centerRange(worksheet, 4, 9, 8, 10);
  centerRange(worksheet, 5, 5, 10, 11);

  worksheet.getCell('E11').value = 'Selección de la muestra';
  worksheet.getCell('E11').font = { name: 'Calibri', size: 11, bold: true, italic: true };

  worksheet.getCell('E12').value = 'Intervalo de selección';
  worksheet.getCell('H12').value = 'k =';
  setFormula(worksheet.getCell('I12'), 'ROUND(I4/J9,0)', stats.interval, '0');
  worksheet.getCell('I12').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: COLORS.yellowFill },
  };
  setThinBorder(worksheet.getCell('I12'));

  worksheet.getCell('E13').value = 'Arranque aleatorio';
  worksheet.getCell('H13').value = 'r =';
  worksheet.getCell('I13').value = stats.start;
  worksheet.getCell('I13').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: COLORS.orangeFill },
  };
  worksheet.getCell('I13').font = { name: 'Calibri', size: 11, bold: true };
  worksheet.getCell('I13').alignment = { horizontal: 'center' };
  setThinBorder(worksheet.getCell('I13'));

  ['E15', 'F15', 'G15', 'H15'].forEach((address) => applyHeaderStyle(worksheet.getCell(address), COLORS.greenFill));
  worksheet.getCell('E15').value = 'N°';
  worksheet.getCell('F15').value = 'Id';
  worksheet.getCell('G15').value = 'RAZÓN SOCIAL';
  worksheet.getCell('H15').value = 'SECTOR';

  data.msSample.forEach((row, index) => {
    const targetRow = index + 16;
    const idFormula = index === 0 ? '$I$13' : `F${targetRow - 1}+$I$12`;

    worksheet.getCell(`E${targetRow}`).value = index + 1;
    worksheet.getCell(`E${targetRow}`).font = { name: 'Calibri', size: 11, color: { argb: COLORS.red } };
    setFormula(worksheet.getCell(`F${targetRow}`), idFormula, row.id, '0');
    setFormula(worksheet.getCell(`G${targetRow}`), `INDEX(MATRIZ!B:B,F${targetRow}+1)`, row.name);
    setFormula(worksheet.getCell(`H${targetRow}`), `INDEX(MATRIZ!C:C,F${targetRow}+1)`, row.sector, '0');
    worksheet.getCell(`E${targetRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`F${targetRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`H${targetRow}`).alignment = { horizontal: 'center' };
  });

  borderRange(worksheet, 15, 15 + data.msSample.length, 5, 8);
}

async function main() {
  const outputPath = process.argv[2] || DEFAULT_OUTPUT_PATH;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(INPUT_PATH);
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.calcProperties.forceFullCalc = true;

  const actividad = workbook.getWorksheet('Actividad');
  const matriz = workbook.getWorksheet('MATRIZ');
  const mas = workbook.getWorksheet('MAS');
  const ms = workbook.getWorksheet('MS');

  const population = [];
  for (let rowNumber = 2; rowNumber <= matriz.rowCount; rowNumber += 1) {
    const id = matriz.getCell(`A${rowNumber}`).value;
    if (id == null || id === '') {
      continue;
    }

    population.push({
      id: Number(id),
      name: asText(matriz.getCell(`B${rowNumber}`).value).trim(),
      sector: Number(matriz.getCell(`C${rowNumber}`).value),
    });
  }

  const N = population.length;
  const conf1 = 0.92;
  const alpha1 = 1 - conf1;
  const z1 = 1.75068607125217;
  const sigma = Math.sqrt(122500);
  const error1 = 100;
  const n1Raw = (N * (z1 ** 2) * (sigma ** 2)) / (((N - 1) * (error1 ** 2)) + ((z1 ** 2) * (sigma ** 2)));
  const n1 = Math.ceil(n1Raw);

  const conf2 = 0.9;
  const z2 = 1.64485362695147;
  const error2 = 0.08;
  const p = 0.85;
  const q = 1 - p;
  const n2Raw = (N * (z2 ** 2) * p * q) / (((N - 1) * (error2 ** 2)) + ((z2 ** 2) * p * q));
  const n2 = Math.ceil(n2Raw);

  const nFinal = Math.max(n1, n2);
  const interval = Math.round(N / nFinal);
  const start = 17;

  const rng = mulberry32(20260420);
  const masSorted = population
    .map((row) => ({
      ...row,
      random: randomInt(rng, 1, 2000),
    }))
    .sort((a, b) => a.random - b.random || a.id - b.id);

  const masSample = masSorted.slice(0, nFinal);
  const msSample = [];
  for (let i = 0; i < nFinal; i += 1) {
    const id = start + (i * interval);
    const row = population[id - 1];
    msSample.push({
      id,
      name: row.name,
      sector: row.sector,
    });
  }

  const stats = {
    N,
    conf1,
    alpha1,
    z1,
    sigma,
    error1,
    n1Raw,
    n1,
    conf2,
    error2,
    z2,
    p,
    q,
    n2Raw,
    n2,
    nFinal,
    interval,
    start,
  };

  buildActividadSheet(actividad, stats);
  buildMasSheet(mas, { masSorted, masSample }, stats);
  buildMsSheet(ms, { population, msSample }, stats);

  await workbook.xlsx.writeFile(outputPath);
  console.log(`Archivo generado en: ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
