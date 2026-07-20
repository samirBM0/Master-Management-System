const XLSX = require('xlsx');

const wb = XLSX.readFile('src/capa_report.xlsx');
const ws = wb.Sheets['CPK Report'];
if (ws) {
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:P3000');
  for (let r = range.s.r; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (cell && cell.f && cell.f.includes('Raw Data')) {
      console.log(`Row ${r+1} (cell ${XLSX.utils.encode_cell({ r, c: 0 })}): ${cell.f}`);
      // Print first 3 and last 3 to be concise
      if (r > 15) {
        console.log("...");
        break;
      }
    }
  }
}
