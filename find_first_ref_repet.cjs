const XLSX = require('xlsx');

const wb = XLSX.readFile('src/repet_report.xlsx');
console.log("Sheet names:", wb.SheetNames);

wb.SheetNames.forEach(sheetName => {
  if (sheetName === 'Raw Data') return;
  const ws = wb.Sheets[sheetName];
  if (!ws) return;
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:P50');
  let found = false;
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= Math.min(range.e.c, 10); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.f && cell.f.includes('Raw Data')) {
        console.log(`Sheet '${sheetName}' Row ${r+1} (cell ${XLSX.utils.encode_cell({ r, c })}): ${cell.f}`);
        found = true;
        break;
      }
    }
    if (found) {
      // Find the first few references and then stop to be concise
      break;
    }
  }
});
