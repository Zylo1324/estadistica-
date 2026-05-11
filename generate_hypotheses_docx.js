const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const baseDocxPath = path.join(__dirname, '_tmp_locked_copy.docx');
const outputDocxPath = path.join(
  __dirname,
  'UNIDAD 2',
  'ejercicios de Hipotesis Estadisticas_resuelto.docx'
);

const exercises = [
  {
    statement:
      'La proporción de empleados hombres que consumen tabaco en una cadena internacional de hoteles es mayor que el de empleados mujeres.',
    development:
      'Se comparan dos proporciones: empleados hombres que consumen tabaco y empleadas mujeres que consumen tabaco. Como el enunciado dice que la proporción de hombres es mayor, la hipótesis alternativa debe ir hacia la derecha.',
    equations: [
      [
        sub('p', 'H'),
        t(': proporción de empleados hombres que consumen tabaco'),
      ],
      [
        sub('p', 'M'),
        t(': proporción de empleadas mujeres que consumen tabaco'),
      ],
      [h(0), t(': '), sub('p', 'H'), t(' ≤ '), sub('p', 'M')],
      [h(1), t(': '), sub('p', 'H'), t(' > '), sub('p', 'M')],
    ],
    response:
      'La prueba es unilateral de cola derecha, porque se desea sostener que la proporción de hombres fumadores es mayor que la de mujeres.',
  },
  {
    statement:
      'En una organización dedicada a la venta de automóviles, el salario promedio de los empleados están por debajo de los 2500 soles.',
    development:
      'La variable de interés es el salario promedio de los empleados. Como el enunciado indica que está por debajo de 2500 soles, la hipótesis alternativa debe plantearse como menor que ese valor.',
    equations: [
      [t('μ: salario promedio de los empleados')],
      [h(0), t(': μ ≥ 2500')],
      [h(1), t(': μ < 2500')],
    ],
    response:
      'La prueba es unilateral de cola izquierda, porque se busca comprobar si el salario promedio es menor que 2500 soles.',
  },
  {
    statement:
      'La proporción de museos a nivel nacional que tiene mayor aceptación del público visitante es mayor al 80%.',
    development:
      'Se trabaja con una sola proporción. El valor de comparación es 80%, es decir, 0,80. Como el enunciado afirma que la proporción es mayor al 80%, la hipótesis alternativa va hacia la derecha.',
    equations: [
      [t('p: proporción de museos con mayor aceptación del público visitante')],
      [h(0), t(': p ≤ 0,80')],
      [h(1), t(': p > 0,80')],
    ],
    response:
      'La prueba es unilateral de cola derecha, porque se desea demostrar que la proporción supera el 80%.',
  },
  {
    statement:
      'La proporción de turistas mujeres mayores de 40 años tiene una diferencia significativa en buenas conductas alimenticias que la proporción de turistas hombres de la misma edad.',
    development:
      'Se comparan dos proporciones: mujeres mayores de 40 años y hombres de la misma edad. Como el enunciado habla de una diferencia significativa, no indica que una proporción sea mayor o menor, sino que son distintas.',
    equations: [
      [
        sub('p', 'M'),
        t(': proporción de turistas mujeres mayores de 40 años con buenas conductas alimenticias'),
      ],
      [
        sub('p', 'H'),
        t(': proporción de turistas hombres mayores de 40 años con buenas conductas alimenticias'),
      ],
      [h(0), t(': '), sub('p', 'M'), t(' = '), sub('p', 'H')],
      [h(1), t(': '), sub('p', 'M'), t(' ≠ '), sub('p', 'H')],
    ],
    response:
      'La prueba es bilateral, porque se busca verificar si existe diferencia significativa entre ambas proporciones.',
  },
  {
    statement:
      'Los restaurantes de la actualidad consumen en promedio un 20% más de energía eléctrica al mes que los de hace veinte años, hace 20 años el consumo era de 1500 Kwatts.',
    development:
      'Primero se calcula el consumo de referencia actual si realmente fuera 20% más que hace veinte años. Como el consumo anterior era 1500 Kwatts, el valor de comparación es 1800 Kwatts.',
    equations: [
      [sub('μ', '0'), t(' = 1500 × 1,20 = 1800')],
      [t('μ: consumo promedio mensual actual de energía eléctrica')],
      [h(0), t(': μ = 1800')],
      [h(1), t(': μ ≠ 1800')],
    ],
    response:
      'La prueba es bilateral, porque el enunciado fija un aumento específico de 20%; por eso se contrasta si el promedio actual corresponde o no a 1800 Kwatts.',
  },
  {
    statement:
      'El gerente de Danvers-Hilton Resort afirma que la cantidad media que gastan los huéspedes en un fin de semana es de $600 o menos. Un miembro del equipo de contadores observó que en los últimos meses habían aumentado tales cantidades. El contador emplea una muestra de cuentas de fin de semana para probar la afirmación del gerente.',
    development:
      'La afirmación del gerente sostiene que el gasto medio es de 600 dólares o menos. El contador sospecha que el gasto ha aumentado, por eso la hipótesis alternativa plantea que el promedio es mayor que 600 dólares.',
    equations: [
      [t('μ: gasto medio de los huéspedes en un fin de semana')],
      [h(0), t(': μ ≤ 600')],
      [h(1), t(': μ > 600')],
    ],
    response:
      'La prueba es unilateral de cola derecha. Se rechazaría la afirmación del gerente si la evidencia muestra que el gasto medio supera los 600 dólares.',
  },
  {
    statement:
      'El gerente de una cadena de restaurante está pensando en un nuevo plan de bonificaciones, con objeto de incrementar el volumen de ventas en sus restaurantes. Al presente, el volumen medio de ventas es de 2950 soles diarios. El gerente desea realizar un estudio para ver si el plan de bonificaciones incrementa el volumen de ventas. Para recolectar los datos toma una muestra de 5 de sus restaurantes que atenderán durante una semana bajo el nuevo plan de bonificaciones.',
    development:
      'La comparación se realiza contra el volumen medio actual de ventas, que es 2950 soles diarios. Como se quiere comprobar si el plan incrementa las ventas, la hipótesis alternativa debe indicar un promedio mayor.',
    equations: [
      [t('μ: volumen medio diario de ventas con el nuevo plan de bonificaciones')],
      [h(0), t(': μ ≤ 2950')],
      [h(1), t(': μ > 2950')],
    ],
    response:
      'La prueba es unilateral de cola derecha, porque se busca comprobar si el nuevo plan incrementa el volumen medio de ventas.',
  },
  {
    statement:
      'Una línea de operación que está diseñada especialmente para atender a los pedidos de una cadena de hoteles, llena empaques de 32 onzas de detergente para lavar. Con periodicidad se selecciona una muestra de los empaques y se pesan para determinar si no se están llenando con un peso mayor o menor del indicado. Si los datos muestrales llevan a la conclusión de que hay exceso o falta de llenado, se suspende la producción y se ajusta al llenado correcto.',
    development:
      'El peso correcto de llenado es 32 onzas. Como el problema menciona exceso o falta de llenado, interesa detectar cualquier desviación respecto a 32 onzas, ya sea por encima o por debajo.',
    equations: [
      [t('μ: peso medio de los empaques de detergente')],
      [h(0), t(': μ = 32')],
      [h(1), t(': μ ≠ 32')],
    ],
    response:
      'La prueba es bilateral, porque se debe detectar tanto exceso como falta de llenado.',
  },
  {
    statement:
      'Debido a los costos y al tiempo de adaptación de la producción, un director de fabricación antes de implantar un nuevo método de fabricación, debe convencer al gerente de que ese nuevo método de fabricación reducirá los costos. El costo medio del actual método de producción es $220 por hora. En un estudio se medirá el costo del nuevo método durante un periodo muestral de producción.',
    development:
      'El costo actual es de 220 dólares por hora. Para justificar el nuevo método, debe mostrarse que el costo medio disminuye; por eso la hipótesis alternativa es menor que 220.',
    equations: [
      [t('μ: costo medio por hora con el nuevo método de fabricación')],
      [h(0), t(': μ ≥ 220')],
      [h(1), t(': μ < 220')],
    ],
    response:
      'La prueba es unilateral de cola izquierda, porque se busca comprobar una reducción del costo medio de producción.',
  },
  {
    statement:
      'Las ventas promedio mensuales de las empresas dedicadas al cultivo y venta de arroz al por mayor en el departamento de Lambayeque es mayor a 12500.',
    development:
      'La variable analizada es el promedio mensual de ventas. Como el enunciado afirma que el promedio es mayor que 12500, la hipótesis alternativa se plantea hacia la derecha.',
    equations: [
      [t('μ: ventas promedio mensuales de las empresas dedicadas al cultivo y venta de arroz')],
      [h(0), t(': μ ≤ 12500')],
      [h(1), t(': μ > 12500')],
    ],
    response:
      'La prueba es unilateral de cola derecha, porque se desea comprobar que las ventas promedio mensuales son mayores a 12500.',
  },
];

function xmlEscape(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function withSpaceAttr(tagName, text) {
  const needsPreserve = /^\s|\s$/.test(text);
  const escaped = xmlEscape(text);
  return needsPreserve
    ? `<${tagName} xml:space="preserve">${escaped}</${tagName}>`
    : `<${tagName}>${escaped}</${tagName}>`;
}

function run(text, options = {}) {
  const { bold = false, italic = false, size = 22 } = options;
  const props = [
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>',
    `<w:sz w:val="${size}"/>`,
    `<w:szCs w:val="${size}"/>`,
  ];
  if (bold) props.push('<w:b/>');
  if (italic) props.push('<w:i/>');
  return `<w:r><w:rPr>${props.join('')}</w:rPr>${withSpaceAttr('w:t', text)}</w:r>`;
}

function paragraph(runs, options = {}) {
  const { align = 'both', before = 0, after = 80, line = 260, indent = 0 } = options;
  const pPr = [
    `<w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/>`,
  ];
  if (align) pPr.push(`<w:jc w:val="${align}"/>`);
  if (indent) pPr.push(`<w:ind w:left="${indent}"/>`);
  return `<w:p><w:pPr>${pPr.join('')}</w:pPr>${runs.join('')}</w:p>`;
}

function blankParagraph(after = 80) {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="${after}" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
}

function t(text) {
  return `<m:r>${withSpaceAttr('m:t', text)}</m:r>`;
}

function wrap(content) {
  return Array.isArray(content) ? content.join('') : content;
}

function sub(baseText, subText) {
  return `<m:sSub><m:e>${t(baseText)}</m:e><m:sub>${t(subText)}</m:sub></m:sSub>`;
}

function h(number) {
  return sub('H', String(number));
}

function equationParagraph(content, options = {}) {
  const { align = 'center', after = 60, before = 0 } = options;
  return [
    '<w:p>',
    '<w:pPr>',
    `<w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/>`,
    `<w:jc w:val="${align}"/>`,
    '</w:pPr>',
    '<m:oMathPara>',
    `<m:oMath>${wrap(content)}</m:oMath>`,
    '</m:oMathPara>',
    '</w:p>',
  ].join('');
}

function exerciseBlock(exercise, index) {
  const parts = [];

  parts.push(
    paragraph([run(`${index + 1}. ${exercise.statement}`, { bold: true })], {
      before: index === 0 ? 80 : 120,
      after: 80,
      align: 'both',
    })
  );

  parts.push(
    paragraph([run('Desarrollo:', { bold: true })], {
      after: 40,
      align: 'left',
    })
  );
  parts.push(paragraph([run(exercise.development)], { after: 70 }));
  exercise.equations.forEach((equation) => {
    parts.push(equationParagraph(equation));
  });

  parts.push(
    paragraph([run('Respuesta:', { bold: true }), run(` ${exercise.response}`)], {
      before: 20,
      after: 100,
    })
  );

  return parts.join('');
}

function buildDocumentXml(originalDocumentXml) {
  const openTagMatch = originalDocumentXml.match(/^([\s\S]*?<w:body>)/);
  const sectPrMatch = originalDocumentXml.match(/(<w:sectPr[\s\S]*<\/w:sectPr>)/);

  if (!openTagMatch || !sectPrMatch) {
    throw new Error('No se pudo recuperar la estructura base del documento.');
  }

  const body = [];
  body.push(
    paragraph([run('Hipótesis Estadísticas', { bold: true, size: 32 })], {
      align: 'center',
      after: 80,
      line: 280,
    })
  );
  body.push(
    paragraph([run('Formule las hipótesis en los siguientes casos:', { bold: true })], {
      align: 'left',
      after: 80,
    })
  );
  exercises.forEach((exercise, index) => {
    body.push(exerciseBlock(exercise, index));
  });
  body.push(blankParagraph(80));

  return `${openTagMatch[1]}${body.join('')}${sectPrMatch[1]}</w:body></w:document>`;
}

function updateStyles(stylesXml) {
  let updated = stylesXml.replace(
    /<w:rFonts[^>]*w:asciiTheme="minorHAnsi"[^>]*\/>/,
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
  );

  updated = updated.replace(
    /<w:sz w:val="\d+"\/>/,
    '<w:sz w:val="22"/>'
  );

  updated = updated.replace(
    /<w:szCs w:val="\d+"\/>/,
    '<w:szCs w:val="22"/>'
  );

  updated = updated.replace(
    /<w:lang w:val="[^"]+" w:eastAsia="[^"]+" w:bidi="[^"]+"\/>/,
    '<w:lang w:val="es-PE" w:eastAsia="es-PE" w:bidi="ar-SA"/>'
  );

  return updated;
}

async function main() {
  const inputBuffer = fs.readFileSync(baseDocxPath);
  const zip = await JSZip.loadAsync(inputBuffer);

  const documentXml = await zip.file('word/document.xml').async('string');
  const stylesXml = await zip.file('word/styles.xml').async('string');

  zip.file('word/document.xml', buildDocumentXml(documentXml));
  zip.file('word/styles.xml', updateStyles(stylesXml));

  const outputBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
  });

  fs.writeFileSync(outputDocxPath, outputBuffer);
  console.log(outputDocxPath);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
