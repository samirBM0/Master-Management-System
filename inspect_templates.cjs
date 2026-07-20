const XLSX = require('xlsx');

function dumpCells(filePath, sheetName) {
  console.log(`\n--- Cells in ${filePath} -> ${sheetName} ---`);
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[sheetName];
  if (!ws) {
    console.log('Sheet not found!');
    return;
  }
  const ref = ws['!ref'] || 'A1:H100';
  const range = XLSX.utils.decode_range(ref);
  for (let r = range.s.r; r <= Math.min(range.e.r, 10); r++) {
    let line = '';
    for (let c = range.s.c; c <= Math.min(range.e.c, 30); c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellRef];
      if (cell && cell.v !== undefined) {
        line += `${cellRef}:[${String(cell.v).substring(0, 40)}]   `;
      }
    }
    if (line.trim()) console.log(`Row ${r+1}: ${line}`);
  }
}

console.log("=== CAPA REPORT RAW DATA ===");
dumpCells('src/capa_report.xlsx', 'Raw Data');

console.log("\n\n=== REPET REPORT RAW DATA ===");
dumpCells('src/repet_report.xlsx', 'Raw Data');
