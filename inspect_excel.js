const ExcelJS = require('exceljs');

async function main() {
  const input = process.argv[2];
  if (!input) {
    throw new Error('Usage: bun inspect_excel.js <xlsx-path>');
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(input);

  for (const worksheet of workbook.worksheets) {
    console.log(`--- ${worksheet.name} ---`);
    console.log(`rows=${worksheet.rowCount} cols=${worksheet.columnCount}`);
    const maxRows = Math.min(worksheet.rowCount, 30);

    for (let rowNumber = 1; rowNumber <= maxRows; rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      const parts = [];
      row.eachCell({ includeEmpty: false }, (cell) => {
        parts.push(`${cell.address}=${JSON.stringify(cell.value)}`);
      });
      if (parts.length > 0) {
        console.log(parts.join(' | '));
      }
    }
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
