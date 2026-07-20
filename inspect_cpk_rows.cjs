const XLSX = require('xlsx');

const wb = XLSX.readFile('src/capa_report.xlsx');
const ws = wb.Sheets['CPK Report'];
if (ws) {
  for (let r = 0; r < 40; r++) {
    let line = '';
    for (let c = 0; c <= 5; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellRef];
      if (cell) {
        line += `${cellRef}:[${String(cell.v).substring(0, 30)}]${cell.f ? ' (formula: ' + cell.f + ')' : ''}   `;
      }
    }
    if (line.trim()) console.log(`Row ${r+1}: ${line}`);
  }
}
