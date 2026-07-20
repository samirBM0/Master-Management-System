import * as XLSX from "xlsx";
import { CapabilityResult, RepeatabilityResult, AirBusEVResult } from "./types";
import { assignGrrMetadata } from "./calculations";

// Read the API token from Vite env (VITE_API_TOKEN in .env / .env.local).
// Vite only exposes variables prefixed with VITE_ to the client.
const API_TOKEN: string | undefined = (import.meta as any).env?.VITE_API_TOKEN;

// Build the Authorization header (Bearer token) if a token is configured.
function authHeaders(): Record<string, string> {
  const h: Record<string, string> = { "Content-Type": "application/json" };
  if (API_TOKEN) h["Authorization"] = `Bearer ${API_TOKEN}`;
  return h;
}

// Download a prepared export by id. Fetches the file as a Blob with the
// Bearer token (browser navigation can't send Authorization headers) and
// triggers the download via a temporary object URL.
async function downloadExport(id: string) {
  const res = await fetch(`/api/export/download/${id}`, {
    method: "GET",
    headers: authHeaders()
  });
  if (!res.ok) throw new Error("Server download returned status " + res.status);

  const blob = await res.blob();
  const contentDisp = res.headers.get("Content-Disposition") || "";
  const nameMatch = contentDisp.match(/filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i);
  const filename = nameMatch
    ? decodeURIComponent(nameMatch[1] || nameMatch[2])
    : "rapport.xlsx";

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Helper function to safely update cell values in place without destroying styles or metadata
function updateCell(ws: XLSX.WorkSheet, r: number, c: number, val: any) {
  const cellRef = XLSX.utils.encode_cell({ r, c });
  if (val === undefined || val === null) {
    delete ws[cellRef];
    return;
  }
  
  const cell: XLSX.CellObject = ws[cellRef] || { t: 's', v: '' };
  cell.v = val;
  if (typeof val === 'number') {
    cell.t = 'n';
  } else if (typeof val === 'boolean') {
    cell.t = 'b';
  } else {
    cell.t = 's';
  }
  
  ws[cellRef] = cell;
  
  const ref = ws['!ref'];
  if (ref) {
    const range = XLSX.utils.decode_range(ref);
    let updated = false;
    if (r < range.s.r) { range.s.r = r; updated = true; }
    if (r > range.e.r) { range.e.r = r; updated = true; }
    if (c < range.s.c) { range.s.c = c; updated = true; }
    if (c > range.e.c) { range.e.c = c; updated = true; }
    if (updated) {
      ws['!ref'] = XLSX.utils.encode_range(range);
    }
  } else {
    ws['!ref'] = XLSX.utils.encode_range({ s: { r, c }, e: { r, c } });
  }
}

export async function exportCapabilityToExcel(
  results: CapabilityResult[],
  rawMeasurements: { [testPointKey: string]: number[] },
  productRef: string,
  testerName: string,
  operatorName: string,
  testName?: string
) {
  try {
    const res = await fetch("/api/export/capability", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({
        results,
        rawMeasurements,
        productRef,
        testerName,
        operatorName,
        testName
      })
    });
    if (!res.ok) throw new Error("Server export returned status " + res.status);
    
    const data = await res.json();
    if (data.success && data.id) {
      await downloadExport(data.id);
    } else {
      throw new Error(data.error || "Failed to prepare export");
    }
  } catch (err) {
    console.error("Server-side capability Excel export failed", err);
    alert("Erreur lors de l'exportation du rapport: " + (err as Error).message);
  }
}

export async function exportRepeatabilityToExcel(
  results: RepeatabilityResult[],
  rawMeasurements: { [testPointKey: string]: number[] },
  productRef: string,
  testerName: string,
  operatorName: string,
  files?: any[],
  testName?: string
) {
  try {
    const res = await fetch("/api/export/repeatability", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({
        results,
        rawMeasurements,
        productRef,
        testerName,
        operatorName,
        files,
        testName
      })
    });
    if (!res.ok) throw new Error("Server export returned status " + res.status);
    
    const data = await res.json();
    if (data.success && data.id) {
      await downloadExport(data.id);
    } else {
      throw new Error(data.error || "Failed to prepare export");
    }
  } catch (err) {
    console.error("Server-side repeatability Excel export failed", err);
    alert("Erreur lors de l'exportation du rapport: " + (err as Error).message);
  }
}

export async function exportAirbusToExcel(
  results: AirBusEVResult[],
  rawMeasurements: { [testPointKey: string]: number[] },
  productRef: string,
  testerName: string,
  operatorName: string
) {
  let wb: XLSX.WorkBook;
  let useTemplate = false;

  try {
    const res = await fetch("/actipa-templates/AIRBUS_MSA_Report.xlsx");
    if (res.ok) {
      const arrayBuffer = await res.arrayBuffer();
      wb = XLSX.read(new Uint8Array(arrayBuffer), { type: "array", cellFormula: true });
      useTemplate = true;
    } else {
      wb = XLSX.utils.book_new();
    }
  } catch {
    wb = XLSX.utils.book_new();
  }

  const headers = [
    "Test Point ID", "Test Step Name", "Unit", "LSL", "USL", "Moyenne", 
    "Ecart-Type (Sigma)", "Range (Max-Min)", "EV Absolue", "EV % (Airbus)", "Statut", "Nombre Mesures"
  ];
  
  let maxMeasurements = 0;
  for (const key in rawMeasurements) {
    maxMeasurements = Math.max(maxMeasurements, rawMeasurements[key].length);
  }
  for (let i = 1; i <= maxMeasurements; i++) {
    headers.push(`Essai ${i}`);
  }

  const rows: any[][] = [
    ["RAPPORT AIRBUS EV%"],
    ["RÉFÉRENCE PRODUIT", productRef],
    ["NOM DU TESTEUR", testerName],
    ["OPÉRATEUR", operatorName],
    ["DATE DU RAPPORT", new Date().toLocaleString()],
    [],
    headers
  ];

  for (const r of results) {
    const key = `${r.testPointId}::${r.testStepName || ""}`;
    const vals = rawMeasurements[key] || [];
    const row = [
      r.testPointId,
      r.testStepName,
      r.unit,
      r.lsl,
      r.usl,
      parseFloat(r.mean.toFixed(6)),
      parseFloat(r.sigma.toFixed(6)),
      parseFloat(r.range.toFixed(6)),
      parseFloat(r.evAbs.toFixed(6)),
      parseFloat((r.evPercent * 100).toFixed(2)) + "%",
      r.status,
      r.valuesCount,
      ...vals
    ];
    rows.push(row);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);

  ws["!cols"] = [
    { wch: 15 }, { wch: 25 }, { wch: 8 }, { wch: 10 }, { wch: 10 }, { wch: 12 },
    { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 12 }
  ];

  if (useTemplate) {
    if (wb.SheetNames.includes("Raw Data")) {
      wb.Sheets["Raw Data"] = ws;
    } else {
      XLSX.utils.book_append_sheet(wb, ws, "Raw Data");
    }
  } else {
    XLSX.utils.book_append_sheet(wb, ws, "Airbus EV Analyse");
  }

  try {
    const wbout = XLSX.write(wb, { type: "base64", bookType: "xlsx" });
    const res = await fetch("/api/export/prepare-generic", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({
        base64Data: wbout,
        filename: `AIRBUS_EV_Rapport_${productRef || "PRODUIT"}.xlsx`,
        contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      })
    });
    if (!res.ok) throw new Error("Server prep returned status " + res.status);
    const data = await res.json();
    if (data.success && data.id) {
      await downloadExport(data.id);
    } else {
      throw new Error(data.error || "Preparation failed");
    }
  } catch (err) {
    console.error("Airbus Excel prep failed, falling back to client-side write", err);
    XLSX.writeFile(wb, `AIRBUS_EV_Rapport_${productRef || "PRODUIT"}.xlsx`);
  }
}
