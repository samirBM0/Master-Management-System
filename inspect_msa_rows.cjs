const XLSX = require('xlsx');

const wb = XLSX.readFile('src/repet_report.xlsx');
const ws = wb.Sheets['MSA Analaysis'];
if (ws) {
  let line = '';
  for (let c = 0; c <= 25; c++) {
    const cellRef = XLSX.utils.encode_cell({ r: 5, c }); // row 6 (0-indexed 5)
    const cell = ws[cellRef];
    if (cell) {
      line += `${cellRef}:[${String(cell.v).substring(0, 30)}]${cell.f ? ' (formula: ' + cell.f + ')' : ''}\n`;
    }
  }
  console.log("MSA Analaysis Row 6 cells:\n" + line);
}
