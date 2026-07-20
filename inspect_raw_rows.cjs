const XLSX = require('xlsx');

const wb = XLSX.readFile('src/capa_report.xlsx');
const ws = wb.Sheets['Raw Data'];
if (ws) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:H100');
  for (let r = 10; r <= 50; r++) {
    let line = '';
    for (let c = 0; c <= 15; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellRef];
      if (cell && cell.v !== undefined) {
        line += `${cellRef}:[${String(cell.v)}] `;
      }
    }
    if (line.trim()) console.log(`Row ${r+1}: ${line}`);
  }
} else {
  console.log('Raw Data sheet not found!');
}
