const XLSX = require('xlsx');

function dumpHeader(filename) {
  console.log(`\n=== Cell layout for Header sheet in ${filename} ===`);
  const wb = XLSX.readFile(filename);
  const ws = wb.Sheets['Header'];
  if (!ws) {
    console.log("No Header sheet found!");
    return;
  }
  const range = XLSX.utils.decode_range(ws['!ref']);
  for (let r = range.s.r; r <= range.e.r; r++) {
    let rowStr = '';
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellRef = XLSX.utils.encode_cell({ r, c });
      const cell = ws[cellRef];
      if (cell && cell.v !== undefined) {
        rowStr += `${cellRef}: "${cell.v}"\t`;
      }
    }
    if (rowStr) {
      console.log(`Row ${r + 1}: ${rowStr}`);
    }
  }
}

dumpHeader('src/repet_report.xlsx');
