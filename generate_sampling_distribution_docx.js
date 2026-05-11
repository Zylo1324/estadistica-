const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const baseDocxPath = path.join(
  __dirname,
  '_tmp_locked_copy.docx'
);

const outputDocxPath = path.join(
  __dirname,
  'Trabajo de distribución muestral para la media muestral y la proporción muestral (1)_resuelto.docx'
);

const Z_ROWS = [
  { z: -1, prob: 0.1587 },
  { z: -0.5, prob: 0.3085 },
  { z: 0, prob: 0.5 },
  { z: 0.5, prob: 0.6915 },
  { z: 1, prob: 0.8413 },
];

const meanMetadata = [
  { mu: 2500, sigma: 600, n: 36 },
  { mu: 320, sigma: 80, n: 49 },
  { mu: 45, sigma: 10, n: 25 },
  { mu: 120, sigma: 30, n: 16 },
  { mu: 1500, sigma: 400, n: 64 },
  { mu: 180, sigma: 50, n: 36 },
  { mu: 2000, sigma: 500, n: 25 },
  { mu: 700, sigma: 150, n: 49 },
  { mu: 25, sigma: 5, n: 100 },
  { mu: 120, sigma: 20, n: 36 },
  { mu: 1800, sigma: 500, n: 25 },
  { mu: 150, sigma: 40, n: 64 },
  { mu: 3, sigma: 0.8, n: 36 },
  { mu: 12, sigma: 2, n: 49 },
  { mu: 25, sigma: 6, n: 25 },
  { mu: 48, sigma: 8, n: 36 },
  { mu: 120, sigma: 30, n: 49 },
  { mu: 5, sigma: 1, n: 64 },
  { mu: 80, sigma: 15, n: 25 },
  { mu: 10000, sigma: 2000, n: 36 },
];

const proportionMetadata = [
  { p: 0.6, n: 100 },
  { p: 0.4, n: 200 },
  { p: 0.3, n: 150 },
  { p: 0.55, n: 100 },
  { p: 0.25, n: 120 },
  { p: 0.7, n: 80 },
  { p: 0.35, n: 200 },
  { p: 0.5, n: 150 },
  { p: 0.45, n: 100 },
  { p: 0.2, n: 80 },
  { p: 0.65, n: 120 },
  { p: 0.15, n: 200 },
  { p: 0.75, n: 100 },
  { p: 0.28, n: 150 },
  { p: 0.52, n: 100 },
  { p: 0.33, n: 120 },
  { p: 0.48, n: 90 },
  { p: 0.22, n: 200 },
  { p: 0.58, n: 150 },
  { p: 0.1, n: 250 },
];

function xmlEscape(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function decodeXml(text) {
  return String(text)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function withSpaceAttr(tagName, text) {
  const needsPreserve = /^\s|\s$/.test(text);
  const escaped = xmlEscape(text);
  return needsPreserve
    ? `<${tagName} xml:space="preserve">${escaped}</${tagName}>`
    : `<${tagName}>${escaped}</${tagName}>`;
}

function formatNumber(value, decimals) {
  const fixed = Number(value).toFixed(decimals);
  return fixed.replace('.', ',');
}

function trimSpanishNumber(value, maxDecimals) {
  const raw = formatNumber(value, maxDecimals);
  return raw
    .replace(/,?0+$/u, '')
    .replace(/,$/u, '');
}

function formatMeanValue(value) {
  if (Math.abs(value - Math.round(value)) < 1e-9) {
    return String(Math.round(value));
  }
  return trimSpanishNumber(value, Math.abs(value) < 10 ? 4 : 2);
}

function formatStandardError(value, kind) {
  if (kind === 'mean') {
    if (Math.abs(value - Math.round(value)) < 1e-9) {
      return String(Math.round(value));
    }
    return trimSpanishNumber(value, 2);
  }
  return formatNumber(value, 4);
}

function formatProbability(value) {
  return formatNumber(value, 4);
}

function formatZ(value) {
  if (Math.abs(value - Math.round(value)) < 1e-9) {
    return String(Math.round(value));
  }
  return String(value).replace('.', ',');
}

function textRun(text, options = {}) {
  const {
    bold = false,
    italic = false,
    size = 22,
  } = options;
  const runProps = [
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>',
    `<w:sz w:val="${size}"/>`,
    `<w:szCs w:val="${size}"/>`,
  ];
  if (bold) runProps.push('<w:b/>');
  if (italic) runProps.push('<w:i/>');
  return `<w:r><w:rPr>${runProps.join('')}</w:rPr>${withSpaceAttr('w:t', text)}</w:r>`;
}

function paragraph(runs, options = {}) {
  const {
    align,
    after = 90,
    before = 0,
    line = 260,
  } = options;
  const pPr = [];
  pPr.push(`<w:spacing w:before="${before}" w:after="${after}" w:line="${line}" w:lineRule="auto"/>`);
  if (align) {
    pPr.push(`<w:jc w:val="${align}"/>`);
  }
  return `<w:p><w:pPr>${pPr.join('')}</w:pPr>${runs.join('')}</w:p>`;
}

function blankParagraph(after = 120) {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="${after}" w:line="240" w:lineRule="auto"/></w:pPr></w:p>`;
}

function mText(text) {
  return `<m:r>${withSpaceAttr('m:t', text)}</m:r>`;
}

function mWrap(content) {
  return Array.isArray(content) ? content.join('') : content;
}

function mSub(base, sub) {
  return `<m:sSub><m:e>${mWrap(base)}</m:e><m:sub>${mWrap(sub)}</m:sub></m:sSub>`;
}

function mSup(base, sup) {
  return `<m:sSup><m:e>${mWrap(base)}</m:e><m:sup>${mWrap(sup)}</m:sup></m:sSup>`;
}

function mFrac(num, den) {
  return `<m:f><m:num>${mWrap(num)}</m:num><m:den>${mWrap(den)}</m:den></m:f>`;
}

function mRad(expr) {
  return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:e>${mWrap(expr)}</m:e></m:rad>`;
}

function mDelim(content, begin = '(', end = ')') {
  return `<m:d><m:dPr><m:begChr m:val="${xmlEscape(begin)}"/><m:endChr m:val="${xmlEscape(end)}"/></m:dPr><m:e>${mWrap(content)}</m:e></m:d>`;
}

function equationParagraph(content) {
  return `<w:p><w:pPr><w:spacing w:before="0" w:after="90" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr><m:oMathPara><m:oMath>${mWrap(content)}</m:oMath></m:oMathPara></w:p>`;
}

function tableCell(text, width, bold = false) {
  return [
    '<w:tc>',
    `<w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>`,
    paragraph([textRun(text, { bold })], { align: 'center', after: 0, before: 0, line: 240 }),
    '</w:tc>',
  ].join('');
}

function table(headers, rows) {
  const grid = [2800, 1800, 2600];
  const tblGrid = `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;
  const tblPr = [
    '<w:tblPr>',
    '<w:tblW w:w="0" w:type="auto"/>',
    '<w:tblBorders>',
    '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>',
    '<w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>',
    '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>',
    '<w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>',
    '<w:insideH w:val="single" w:sz="8" w:space="0" w:color="000000"/>',
    '<w:insideV w:val="single" w:sz="8" w:space="0" w:color="000000"/>',
    '</w:tblBorders>',
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>',
    '</w:tblPr>',
  ].join('');

  const headerRow = `<w:tr>${headers.map((header, index) => tableCell(header, grid[index], true)).join('')}</w:tr>`;
  const dataRows = rows
    .map((row) => `<w:tr>${row.map((cell, index) => tableCell(cell, grid[index], false)).join('')}</w:tr>`)
    .join('');

  return `<w:tbl>${tblPr}${tblGrid}${headerRow}${dataRows}</w:tbl>`;
}

function extractOriginalTexts(documentXml) {
  return [...documentXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((match) => decodeXml(match[1]));
}

function meanResponse(index, muText, seText, n) {
  const variants = [
    ` La media muestral queda aproximadamente normal con centro en ${muText} y error estándar ${seText}. Por eso, los promedios más cercanos a ${muText} son los que mejor representan el comportamiento esperado de la muestra.`,
    ` La distribución de la media muestral está centrada en ${muText} y su error estándar es ${seText}. En una muestra de tamaño ${n}, los valores alrededor de esa media son los más razonables para describir el promedio observado.`,
    ` El promedio muestral sigue una distribución aproximadamente normal con media ${muText} y error estándar ${seText}. En consecuencia, los valores situados cerca de ${muText} son los más habituales dentro del comportamiento muestral.`,
  ];
  return variants[index % variants.length];
}

function proportionResponse(index, pText, seText) {
  const variants = [
    ` La proporción muestral se distribuye aproximadamente normal con media ${pText} y error estándar ${seText}. Por eso, las proporciones cercanas a ${pText} son las que mejor representan el comportamiento esperado.`,
    ` La distribución de la proporción muestral queda centrada en ${pText} y presenta un error estándar de ${seText}. En consecuencia, los valores próximos a esa proporción poblacional son los más consistentes con la muestra.`,
    ` La proporción muestral sigue aproximadamente una normal con centro en ${pText} y dispersión ${seText}. Así, las proporciones vecinas a ${pText} son las más naturales dentro del comportamiento muestral.`,
  ];
  return variants[index % variants.length];
}

function buildMeanExercise(index, statement, meta) {
  const sqrtN = Math.sqrt(meta.n);
  const se = meta.sigma / sqrtN;
  const muText = formatMeanValue(meta.mu);
  const sigmaText = formatMeanValue(meta.sigma);
  const sqrtNText = formatMeanValue(sqrtN);
  const seText = formatStandardError(se, 'mean');

  const referenceValues = Z_ROWS.map((row) => meta.mu + row.z * se);
  const rows = referenceValues.map((value, idx) => [
    formatMeanValue(value),
    formatZ(Z_ROWS[idx].z),
    formatProbability(Z_ROWS[idx].prob),
  ]);

  const parts = [];
  parts.push(paragraph([textRun(`${index + 1}. ${statement}`)], { align: 'both', after: 100 }));
  parts.push(paragraph([textRun('Desarrollo:', { bold: true })], { after: 50 }));
  parts.push(
    paragraph(
      [
        textRun(
          `Tomo como datos una media poblacional de ${muText}, una desviación estándar de ${sigmaText} y una muestra de ${meta.n} observaciones.`
        ),
      ],
      { align: 'both', after: 70 }
    )
  );
  parts.push(paragraph([textRun('La media de la distribución muestral coincide con la media poblacional:')], { after: 50 }));
  parts.push(
    equationParagraph([
      mSub(mText('μ'), mText('X̄')),
      mText(' = '),
      mText('μ'),
      mText(' = '),
      mText(muText),
    ])
  );
  parts.push(paragraph([textRun('Luego calculo el error estándar de la media muestral:')], { after: 50 }));
  parts.push(
    equationParagraph([
      mSub(mText('σ'), mText('X̄')),
      mText(' = '),
      mFrac(mText(sigmaText), mRad(mText(String(meta.n)))),
      mText(' = '),
      mFrac(mText(sigmaText), mText(sqrtNText)),
      mText(' = '),
      mText(seText),
    ])
  );
  parts.push(paragraph([textRun('Con ese resultado, la distribución muestral queda:')], { after: 50 }));
  parts.push(
    equationParagraph([
      mText('X̄'),
      mText(' ∼ N'),
      mDelim([
        mText(muText),
        mText('; '),
        mSup(mText(seText), mText('2')),
      ]),
    ])
  );
  parts.push(
    paragraph(
      [
        textRun(
          'Para analizar su comportamiento, tomo cinco valores ubicados a una y a media desviación estándar alrededor de la media.'
        ),
      ],
      { align: 'both', after: 50 }
    )
  );
  parts.push(
    equationParagraph([
      mText('z = '),
      mFrac(
        [
          mText('X̄'),
          mText(' - '),
          mSub(mText('μ'), mText('X̄')),
        ],
        mSub(mText('σ'), mText('X̄'))
      ),
    ])
  );
  parts.push(
    table(
      ['Valor de x̄', 'Z', 'Probabilidad acumulada'],
      rows
    )
  );
  parts.push(
    paragraph(
      [
        textRun('Respuesta:', { bold: true }),
        textRun(meanResponse(index, muText, seText, meta.n)),
      ],
      { align: 'both', before: 40, after: 130 }
    )
  );
  parts.push(blankParagraph(130));
  return parts.join('');
}

function buildProportionExercise(index, statement, meta) {
  const q = 1 - meta.p;
  const se = Math.sqrt((meta.p * q) / meta.n);
  const pText = formatNumber(meta.p, 2);
  const qText = formatNumber(q, 2);
  const seText = formatStandardError(se, 'proportion');
  const npText = trimSpanishNumber(meta.p * meta.n, 2);
  const nqText = trimSpanishNumber(q * meta.n, 2);

  const referenceValues = Z_ROWS.map((row) => meta.p + row.z * se);
  const rows = referenceValues.map((value, idx) => [
    formatNumber(value, 4),
    formatZ(Z_ROWS[idx].z),
    formatProbability(Z_ROWS[idx].prob),
  ]);

  const parts = [];
  parts.push(paragraph([textRun(`${index + 1}. ${statement}`)], { align: 'both', after: 100 }));
  parts.push(paragraph([textRun('Desarrollo:', { bold: true })], { after: 50 }));
  parts.push(
    paragraph(
      [
        textRun(
          `La proporción poblacional es ${pText}, su complementaria es ${qText} y el tamaño de muestra es ${meta.n}. Además, n·p = ${npText} y n·q = ${nqText}, por lo que la aproximación normal es válida.`
        ),
      ],
      { align: 'both', after: 70 }
    )
  );
  parts.push(paragraph([textRun('La media de la distribución muestral de la proporción es:')], { after: 50 }));
  parts.push(
    equationParagraph([
      mSub(mText('μ'), mText('p̂')),
      mText(' = '),
      mText('p'),
      mText(' = '),
      mText(pText),
    ])
  );
  parts.push(paragraph([textRun('Ahora calculo el error estándar de la proporción muestral:')], { after: 50 }));
  parts.push(
    equationParagraph([
      mSub(mText('σ'), mText('p̂')),
      mText(' = '),
      mRad(
        mFrac(
          [
            mText(pText),
            mText('·'),
            mText(qText),
          ],
          mText(String(meta.n))
        )
      ),
      mText(' = '),
      mText(seText),
    ])
  );
  parts.push(paragraph([textRun('Entonces la distribución muestral aproximada queda:')], { after: 50 }));
  parts.push(
    equationParagraph([
      mText('p̂'),
      mText(' ∼ N'),
      mDelim([
        mText(pText),
        mText('; '),
        mSup(mText(seText), mText('2')),
      ]),
    ])
  );
  parts.push(
    paragraph(
      [
        textRun(
          'Para analizar la proporción muestral, tomo cinco valores de referencia alrededor de la proporción poblacional.'
        ),
      ],
      { align: 'both', after: 50 }
    )
  );
  parts.push(
    equationParagraph([
      mText('z = '),
      mFrac(
        [
          mText('p̂'),
          mText(' - '),
          mText('p'),
        ],
        mSub(mText('σ'), mText('p̂'))
      ),
    ])
  );
  parts.push(
    table(
      ['Valor de p̂', 'Z', 'Probabilidad acumulada'],
      rows
    )
  );
  parts.push(
    paragraph(
      [
        textRun('Respuesta:', { bold: true }),
        textRun(proportionResponse(index, pText, seText)),
      ],
      { align: 'both', before: 40, after: 130 }
    )
  );
  parts.push(blankParagraph(130));
  return parts.join('');
}

function buildDocumentXml(originalDocumentXml) {
  const allTexts = extractOriginalTexts(originalDocumentXml);
  const title = allTexts[0];
  const statements = allTexts.slice(1);

  if (statements.length !== 40) {
    throw new Error(`Se esperaban 40 ejercicios y se encontraron ${statements.length}.`);
  }

  const openTagMatch = originalDocumentXml.match(/^([\s\S]*?<w:body>)/);
  const sectPrMatch = originalDocumentXml.match(/(<w:sectPr[\s\S]*<\/w:sectPr>)/);

  if (!openTagMatch || !sectPrMatch) {
    throw new Error('No se pudo recuperar la estructura base del documento.');
  }

  const body = [];
  body.push(
    paragraph(
      [textRun(title, { bold: true, size: 28 })],
      { align: 'center', after: 160, line: 280 }
    )
  );
  body.push(blankParagraph(140));

  statements.forEach((statement, index) => {
    if (index < 20) {
      body.push(buildMeanExercise(index, statement, meanMetadata[index]));
    } else {
      body.push(buildProportionExercise(index, statement, proportionMetadata[index - 20]));
    }
  });

  return `${openTagMatch[1]}${body.join('')}${sectPrMatch[1]}</w:body></w:document>`;
}

function updateStyles(stylesXml) {
  let updated = stylesXml.replace(
    /<w:rFonts[^>]*w:asciiTheme="minorHAnsi"[^>]*\/>/,
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
  );

  updated = updated.replace(
    /<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"\/>/,
    '<w:lang w:val="es-PE" w:eastAsia="es-PE" w:bidi="ar-SA"/>'
  );

  return updated;
}

async function main() {
  const inputBuffer = fs.readFileSync(baseDocxPath);
  const zip = await JSZip.loadAsync(inputBuffer);

  const documentXml = await zip.file('word/document.xml').async('string');
  const stylesXml = await zip.file('word/styles.xml').async('string');

  const newDocumentXml = buildDocumentXml(documentXml);
  const newStylesXml = updateStyles(stylesXml);

  zip.file('word/document.xml', newDocumentXml);
  zip.file('word/styles.xml', newStylesXml);

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
