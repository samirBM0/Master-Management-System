import express from "express";
import path from "path";
import fs from "fs";
import crypto from "crypto";
import { createServer as createViteServer } from "vite";
import { GoogleGenAI, Type } from "@google/genai";
import dotenv from "dotenv";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { INITIAL_MASTER_DATA } from "./src/data";
import { db, countUsers, createUser, getUserByMatricule, listTechnicians, deleteTechnician, revokeToken, purgeExpiredRevocations } from "./src/db";
import { requireAuth, requireRole, signToken, hashPassword, comparePassword } from "./src/auth";

dotenv.config();

// Unified resolver for xlsx module in CJS/ESM/tsx contexts
const xlsxModule = XLSX.readFile ? XLSX : ((XLSX as any).default as any);

// Strip Excel/CSV formula-injection prefixes before writing to a cell
function safeCell(value: unknown): string {
  const s = value == null ? "" : String(value).trim();
  return s.replace(/^[=+\-@\t\r]/, "'$&");
}

// Sanitize filenames to remove path/shell metacharacters
function sanitizeFilename(name: unknown): string {
  return (name == null ? "" : String(name))
    .replace(/[^A-Za-z0-9 _.\-À-ÿ]/g, "_")
    .slice(0, 100);
}

function formatExcelDate(val: any): string {
  if (val === undefined || val === null) return "";
  const trimmed = String(val).trim();
  if (!trimmed) return "";
  
  // Check if it's an Excel serial date number
  const num = Number(trimmed);
  if (!isNaN(num) && num > 30000 && num < 60000) {
    const date = new Date(Math.round((num - 25569) * 86400 * 1000));
    const day = String(date.getUTCDate()).padStart(2, "0");
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    const year = date.getUTCFullYear();
    return `${day}/${month}/${year}`;
  }
  return trimmed;
}

const app = express();
const PORT = 3005;

// Security: per-route body size + basic rate limiting on the API surface
app.use("/api/", (req, res, next) => {
  // express.json/urlencoded are applied as app.use("/api/"), so Express strips the
  // "/api/" mount prefix from req.path (e.g. "/export/capability"). Build the path
  // against both the stripped form and the original URL so the whitelist resolves
  // correctly regardless of how the middleware is mounted.
  const reqPath = (req.path || "").replace(/\/+$/, "");
  const reqFull = (req.originalUrl || req.url || "").split("?")[0].replace(/\/+$/, "");
  const whitelist = ["/export/prepare-generic", "/export/capability", "/export/repeatability",
    "/api/export/prepare-generic", "/api/export/capability", "/api/export/repeatability"];
  const limit = (whitelist.includes(reqPath) || whitelist.includes(reqFull)) ? "20mb" : "2mb";
  const jsonOpts = express.json({ limit });
  jsonOpts(req, res, (err1: any) => {
    if (err1) {
      // entity.too.large -> 413, anything else (e.g. JSON syntax) -> 400 Bad Request
      if (err1.status === 413 || err1.type === "entity.too.large") {
        return res.status(413).json({ error: "Payload trop volumineux." });
      }
      return res.status(400).json({ error: "Corps de requête invalide: " + (err1.message || "JSON invalide") });
    }
    const urlencodedOpts = express.urlencoded({ limit, extended: true });
    urlencodedOpts(req, res, (err2: any) => {
      if (err2) {
        if (err2.status === 413 || err2.type === "entity.too.large") {
          return res.status(413).json({ error: "Payload trop volumineux." });
        }
        return res.status(400).json({ error: "Corps de requête invalide: " + (err2.message || "données invalide") });
      }
      next();
    });
  });
});
app.use("/api/", requireAuth);

// Initialize Gemini SDK with telemetry headers and process.env.GEMINI_API_KEY
const getAiClient = () => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY environment variable is required but missing.");
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      }
    }
  });
};

// API Endpoint for AI-powered Capability calculation
app.post("/api/gemini/calculate-capability", requireAuth, async (req, res) => {
  try {
    const { mesuresBrutes, lsl, usl, nominal, estSC } = req.body;

    // --- Input validation / sanitization (prevents prompt injection) ---
    if (!Array.isArray(mesuresBrutes) || mesuresBrutes.length === 0 || mesuresBrutes.length > 1000) {
      return res.status(400).json({ error: "Le tableau de mesures 'mesuresBrutes' est requis (1..1000 nombres)." });
    }
    const clean = mesuresBrutes
      .map((v: unknown) => Number(v))
      .filter((v: number) => Number.isFinite(v) && Math.abs(v) < 1e15);
    if (clean.length === 0) {
      return res.status(400).json({ error: "Aucune mesure numérique valide fournie." });
    }
    const num = (v: unknown) => { const n = Number(v); return Number.isFinite(n) ? n : NaN; };
    const cLsl = num(lsl), cUsl = num(usl), cNom = num(nominal);
    if ([cLsl, cUsl, cNom].some((n) => Number.isNaN(n))) {
      return res.status(400).json({ error: "lsl/usl/nominal doivent être numériques." });
    }

    const ai = getAiClient();
    
    // Build prompt from validated numbers only (no raw user strings injected)
    const thresholdCpk = estSC ? 1.67 : 1.33;
    
    const prompt = `
Tu es un ingénieur métrologue expert en capabilité statistique des procédés industriels (normes ISO/TS).
Analyse les relevés de mesures brutes suivants d'un banc de test :
- Relevés de mesures brutes (mesuresBrutes): [${clean.join(", ")}]
- LSL (Lower Spec Limit / Limite Inférieure): ${cLsl}
- USL (Upper Spec Limit / Limite Supérieure): ${cUsl}
- Nominal (Valeur cible): ${cNom}
- Est-ce une Caractéristique Spéciale (SC) critique ?: ${estSC ? "OUI (Seuil Cpk exigé ≥ 1.67)" : "NON (Seuil Cpk exigé ≥ 1.33)"}

Consignes impératives pour les calculs statistiques standard :
1. Calcule la moyenne des relevés.
2. Calcule l'écart-type d'échantillon (formule avec diviseur N-1).
3. Calcule l'étendue R (Max - Min).
4. Calcule la répétabilité EV (6 * écart-type).
5. Calcule l'indice de capabilité Cp = (USL - LSL) / (6 * écart-type).
6. Calcule l'indice de capabilité ajusté Cpk = min((USL - moyenne) / (3 * écart-type), (moyenne - LSL) / (3 * écart-type)).

Consignes pour la détection des valeurs aberrantes (Outliers) :
1. Analyse la série de mesures et identifie les valeurs suspectes ou aberrantes (outliers) en utilisant une méthode statistique (par exemple l'écart à plus de 3 écarts-types ou la méthode de Tukey d'écart interquartile).
2. Fournis un diagnostic métrologique et une explication qualitative concise en français dans "explication_diagnostic" expliquant le statut du banc, les valeurs anormales éventuelles trouvées et des conseils d'ajustement ou d'étalonnage.
3. Si des valeurs aberrantes sont éliminées, recalcule les métriques correspondantes sur l'échantillon nettoyé et renvoie-les dans "details_metriques". Si aucune valeur aberrante n'est détectée, renvoie les calculs standard.
4. L'évaluation générale "statut_evaluation" doit être "CONFORME" si l'ajuste_cpk final recalculé est supérieur ou égal au seuil exigé (${thresholdCpk}), ou "MARGINAL" si l'ajuste_cpk est proche (à moins de 0.3 du seuil), ou sinon "INSUFFISANT / REJETÉ".

Renvoie un objet JSON respectant exactement le schéma suivant :
{
  "statut_evaluation": "string (CONFORME, MARGINAL, ou INSUFFISANT / REJETÉ)",
  "explication_diagnostic": "string (explication qualitative rédigée en français)",
  "outliers_detectes": [nombre(s) identifié(s) comme aberration],
  "details_metriques": {
    "potentiel_cp": nombre (Cp calculé/ajusté),
    "ajuste_cpk": nombre (Cpk calculé/ajusté),
    "moyenne_av": nombre (moyenne de l'échantillon nettoyé),
    "ecart_type_std": nombre (écart-type de l'échantillon nettoyé),
    "etendue_r": nombre (étendue de l'échantillon nettoyé),
    "repetabilite_ev": nombre (répétabilité de l'échantillon nettoyé)
  },
  "mesures_nettoyees": [tableau des valeurs restantes après filtrage des aberrations]
}
`;

    const response = await ai.models.generateContent({
      model: 'gemini-3.5-flash',
      contents: prompt,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.OBJECT,
          properties: {
            statut_evaluation: { type: Type.STRING },
            explication_diagnostic: { type: Type.STRING },
            outliers_detectes: {
              type: Type.ARRAY,
              items: { type: Type.NUMBER }
            },
            details_metriques: {
              type: Type.OBJECT,
              properties: {
                potentiel_cp: { type: Type.NUMBER },
                ajuste_cpk: { type: Type.NUMBER },
                moyenne_av: { type: Type.NUMBER },
                ecart_type_std: { type: Type.NUMBER },
                etendue_r: { type: Type.NUMBER },
                repetabilite_ev: { type: Type.NUMBER }
              },
              required: ["potentiel_cp", "ajuste_cpk", "moyenne_av", "ecart_type_std", "etendue_r", "repetabilite_ev"]
            },
            mesures_nettoyees: {
              type: Type.ARRAY,
              items: { type: Type.NUMBER }
            }
          },
          required: ["statut_evaluation", "explication_diagnostic", "outliers_detectes", "details_metriques", "mesures_nettoyees"]
        }
      }
    });

    const resultText = response.text;
    if (!resultText) {
      throw new Error("Réponse vide de la part du modèle d'intelligence artificielle.");
    }

    const resultJson = JSON.parse(resultText);
    res.json({ success: true, result: resultJson });
  } catch (error: any) {
    console.error("Error in AI capability calculation:", error);
    res.status(500).json({ success: false, error: error.message || "Erreur interne lors du traitement par l'IA" });
  }
});

// Path to the master excel file
// Data directory: resolve relative to the server file itself (via __dirname_shim)
// rather than process.cwd(), so file paths stay correct inside the Docker container
// where the working directory is the image WORKDIR (/app). This keeps the Excel
// database and technicians.json accessible regardless of the current working dir.
// Source-data directory (Excel master DB + report templates + assets).
// Resolved against the working directory so it maps to /app/src inside Docker.
const DATA_DIR = path.join(process.cwd(), "src");
const EXCEL_PATH = process.env.EXCEL_PATH || path.join(DATA_DIR, "FR 509-B Suivi pièces master.xlsx");
const EXCEL_WRITE_PATH = EXCEL_PATH;

// Helper to update /src/data.ts file with latest master items to keep it in sync
function updateDataTsFile(items: any[]) {
  try {
    const filePath = path.join(process.cwd(), "src", "data.ts");
    const content = `import { MasterItem } from "./types";\n\nexport const INITIAL_MASTER_DATA: MasterItem[] = ${JSON.stringify(items, null, 2)};\n`;
    
    if (fs.existsSync(filePath)) {
      const existing = fs.readFileSync(filePath, "utf-8");
      if (existing.trim() === content.trim()) {
        console.log("/src/data.ts is already perfectly synchronized with the Excel file. Skipping rewrite.");
        return true;
      }
    }

    fs.writeFileSync(filePath, content, "utf-8");
    console.log(`Successfully updated /src/data.ts file with ${items.length} items from Excel.`);
    return true;
  } catch (error) {
    console.error("Failed to write to /src/data.ts:", error);
    return false;
  }
}

// Load master items from Excel, or fallback to INITIAL_MASTER_DATA
function loadMastersFromExcel() {
  try {
    if (!fs.existsSync(EXCEL_PATH)) {
      console.log(`Excel file not found at ${EXCEL_PATH}. Falling back to initial data.`);
      return INITIAL_MASTER_DATA;
    }

    const workbook = xlsxModule.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = xlsxModule.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

    // Find header row containing TESTEURS
    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      if (row && row[0] && String(row[0]).trim().toUpperCase() === "TESTEURS") {
        headerIdx = i;
        break;
      }
    }

    if (headerIdx === -1) {
      console.log('Could not find headers row with "TESTEURS". Falling back to initial data.');
      return INITIAL_MASTER_DATA;
    }

    const parsed: any[] = [];
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const idMaster = String(row[1] || "").trim();
      if (!idMaster) continue; // Skip empty rows or rows without an ID

      parsed.push({
        id: idMaster,
        testeur: String(row[0] || "").trim(),
        idMaster,
        refProduitMaster: String(row[2] || "").trim(),
        numSerieProduitMaster: String(row[3] || "").trim(),
        refCarteMaster: String(row[4] || "").trim(),
        numSerieCarteMaster: String(row[5] || "").trim(),
        dateCreation: formatExcelDate(row[6]),
        commentaire1: String(row[7] || "").trim(),
        statut: String(row[8] || "").trim() || "Active",
        commentaire2: String(row[9] || "").trim(),
        verif: String(row[10] || "").trim() || "OK"
      });
    }

    if (parsed.length > 0) {
      console.log(`Successfully parsed ${parsed.length} master items from Excel.`);
      // Sync the parsed Excel items to src/data.ts automatically on load
      updateDataTsFile(parsed);
      return parsed;
    } else {
      console.log("No valid items parsed from Excel, using initial data.");
      return INITIAL_MASTER_DATA;
    }
  } catch (error) {
    console.error("Error loading masters from Excel:", error);
    return INITIAL_MASTER_DATA;
  }
}

// Save master items back to Excel
function saveMastersToExcel(items: any[]) {
  try {
    const headers = [
      "TESTEURS",
      "IdentificationMaster",
      "Réf Produit Master",
      "N° série Produit Master",
      "Réf Carte Master ",
      "N° série Carte Master",
      "Date de Création",
      "Commentaire ",
      "Statut de la pce Master",
      "Commentaire",
      "verif"
    ];

    const data = [
      ["", "Suivi des pièces masters", "", "", "", "", "", "", "", "FR 509"],
      ["M-à-j par: BEN MANSOUR Samir", "", "", "", "", "", "", "", "", "", "", "", ""],
      ["Date: " + new Date().toLocaleDateString('fr-FR'), "", "", "", "", "", "", "", "", "", "", "", ""],
      ["Verion: B", "", "", "", "", "", "", "", "", "", "", "", ""],
      headers,
      ...items.map(item => [
        item.testeur || "",
        item.idMaster || "",
        item.refProduitMaster || "",
        item.numSerieProduitMaster || "",
        item.refCarteMaster || "",
        item.numSerieCarteMaster || "",
        item.dateCreation || "",
        item.commentaire1 || "",
        item.statut || "Active",
        item.commentaire2 || "",
        item.verif || "OK"
      ])
    ];

    const worksheet = xlsxModule.utils.aoa_to_sheet(data);
    const workbook = xlsxModule.utils.book_new();
    xlsxModule.utils.book_append_sheet(workbook, worksheet, "MASTER FCT");

    xlsxModule.writeFile(workbook, EXCEL_WRITE_PATH);
    console.log(`Successfully wrote ${items.length} master items back to Excel at ${EXCEL_WRITE_PATH}.`);
    
    // Re-read the freshly written Excel file to guarantee src/data.ts reflects
    // EXACTLY what is now persisted on disk (true resync from the source of truth).
    const reloaded = loadMastersFromExcel();
    updateDataTsFile(reloaded.length ? reloaded : items);
    return true;
  } catch (error) {
    console.error("Error saving masters to Excel:", error);
    return false;
  }
}

// Resync the in-memory data source (src/data.ts) from the authoritative Excel
// file so the UI picks up the latest persisted changes instantly.
function resyncMastersToDataSource(): any[] {
  const reloaded = loadMastersFromExcel();
  updateDataTsFile(reloaded);
  return reloaded;
}

// Serve Excel report templates
app.get("/actipa-templates/capa_report.xlsx", (req, res) => {
  res.sendFile(path.join(DATA_DIR, "capa_report.xlsx"));
});

app.get("/actipa-templates/repet_report.xlsx", async (req, res) => {
  try {
    const candidates = [
      path.resolve(process.cwd(), 'repet_report.xlsx'),
      path.join(DATA_DIR, 'repet_report.xlsx')
    ];
    const templatePath = candidates.find(p => fs.existsSync(p));
    if (!templatePath) {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet("Header");
      wb.addWorksheet("Raw Data");
      const headerWs = wb.getWorksheet("Header");
      if (headerWs) {
        headerWs.getCell("D10").value = "";
        headerWs.getCell("D11").value = "";
        headerWs.getCell("C21").value = "";
      }
      const buffer = await wb.xlsx.writeBuffer();
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent("repet_report.xlsx")}`);
      res.send(buffer);
      return;
    }
    res.sendFile(templatePath);
  } catch (err: any) {
    console.error("Error serving repet_report.xlsx:", err);
    res.status(500).json({ error: "L'exportation a échoué côté serveur.", details: err.message });
  }
});

// Serve Actipa image assets from /src
app.get("/actipa-assets/:filename", (req, res) => {
  const allowed = [
    "LOGO_ACTIA.png",
    "LOGO_ACTIA.jpg",
    "ICON.PNG",
    "ICON.jpg",
    "appicon.ico",
    "legendeCPK2.png",
    "legendeGRR2.png",
    "legendeK2.png",
    "legendeNDC2.PNG"
  ];
  const filename = req.params.filename;
  if (allowed.includes(filename)) {
    res.sendFile(path.join(DATA_DIR, filename));
  } else {
    res.status(404).send("Not found");
  }
});

// Helper for GRR file assignments on backend
function assignGrrMetadataBackend(filename: string, index: number, total: number) {
  const lower = filename.toLowerCase();
  
  let operatorIndex = 0;
  if (lower.includes("op2") || lower.includes("oper2") || lower.includes("operateur2") || lower.includes("op_2") || lower.includes("operator2") || lower.includes("tech_b") || lower.includes("techb") || lower.includes("o2")) {
    operatorIndex = 1;
  } else if (lower.includes("op1") || lower.includes("oper1") || lower.includes("operateur1") || lower.includes("op_1") || lower.includes("operator1") || lower.includes("tech_a") || lower.includes("techa") || lower.includes("o1")) {
    operatorIndex = 0;
  } else {
    operatorIndex = index < (total / 2) ? 0 : 1;
  }

  let partIndex = 0;
  if (lower.includes("p2") || lower.includes("part2") || lower.includes("piece2") || lower.includes("part_2") || lower.includes("piece_2") || lower.includes("partb") || lower.includes("part b")) {
    partIndex = 1;
  } else if (lower.includes("p1") || lower.includes("part1") || lower.includes("piece1") || lower.includes("part_1") || lower.includes("piece_1") || lower.includes("parta") || lower.includes("part a")) {
    partIndex = 0;
  } else {
    const operatorFilesCount = Math.ceil(total / 2);
    const localIndex = index < operatorFilesCount ? index : (index - operatorFilesCount);
    partIndex = localIndex < (operatorFilesCount / 2) ? 0 : 1;
  }

  let trialIndex = 0;
  if (lower.includes("t3") || lower.includes("trial3") || lower.includes("essai3") || lower.includes("run3") || lower.includes("run_3")) {
    trialIndex = 2;
  } else if (lower.includes("t2") || lower.includes("trial2") || lower.includes("essai2") || lower.includes("run2") || lower.includes("run_2")) {
    trialIndex = 1;
  } else if (lower.includes("t1") || lower.includes("trial1") || lower.includes("essai1") || lower.includes("run1") || lower.includes("run_1")) {
    trialIndex = 0;
  } else {
    trialIndex = index % 3;
  }

  return {
    filename,
    operatorIndex,
    partIndex,
    trialIndex
  };
}

// Map to temporarily store prepared Excel exports for sandboxed iframe-safe downloading
const preparedExports = new Map<string, {
  buffer: Buffer;
  filename: string;
  contentType: string;
  createdAt: number;
}>();

// Generate an unguessable 256-bit token (fixes V1 IDOR + V6 predictable ids)
function newExportToken(): string {
  return crypto.randomBytes(32).toString("hex");
}

// API Endpoint to export Capability to Excel using exceljs
app.post("/api/export/capability", requireAuth, async (req, res) => {
  try {
    const { results, rawMeasurements, productRef, testerName, operatorName, isScMode, testName } = req.body;
    
    const templatePath = path.join(DATA_DIR, "capa_report.xlsx");
    let templateFile: string | null = null;
    try {
      if (fs.statSync(templatePath).isFile()) {
        templateFile = templatePath;
      }
    } catch {
      templateFile = null;
    }

    const workbook = new ExcelJS.Workbook();
    if (templateFile) {
      await workbook.xlsx.readFile(templateFile);
    } else {
      console.warn(`capa_report.xlsx not found at expected paths. Creating fallback workbook dynamically.`);
      workbook.addWorksheet("Header");
      workbook.addWorksheet("Raw Data");
    }
    
    const wsHeader = workbook.getWorksheet("Header");
    const wsRaw = workbook.getWorksheet("Raw Data");
    
    if (wsHeader) {
      wsHeader.getCell("D10").value = testerName + " - " + productRef;
      wsHeader.getCell("D11").value = productRef;
      wsHeader.getCell("D12").value = "V1.0";
      wsHeader.getCell("C17").value = testerName;
      wsHeader.getCell("E17").value = testName || (testerName + " Bench");
      wsHeader.getCell("D39").value = operatorName;
      wsHeader.getCell("F39").value = new Date().toLocaleDateString('fr-FR');
      
      const allOk = results.every((r: any) => r.status === "OK");
      wsHeader.getCell("C21").value = allOk ? "OK" : "NOK";
    }
    
    if (wsRaw) {
      // Clear rows 13 to 200 in Raw Data sheet to prevent stale data remnants
      for (let r = 13; r <= 200; r++) {
        const row = wsRaw.getRow(r);
        row.eachCell((cell) => {
          cell.value = null;
        });
      }
      
      // Populate test data
      for (let idx = 0; idx < results.length; idx++) {
        const r = results[idx];
        const key = `${r.testPointId}::${r.testStepName || ""}`;
        const vals = rawMeasurements[key] || [];
        const rowIndex = 13 + idx;
        const row = wsRaw.getRow(rowIndex);
        
        row.getCell(2).value = r.testPointId; // Col B
        row.getCell(3).value = r.testStepName; // Col C
        row.getCell(4).value = r.lsl; // Col D
        row.getCell(5).value = r.usl; // Col E
        row.getCell(6).value = r.unit || ""; // Col F
        
        for (let valIdx = 0; valIdx < vals.length; valIdx++) {
          row.getCell(7 + valIdx).value = vals[valIdx]; // Col G onwards
        }
        row.commit();
      }
    }
    
    const buffer = await workbook.xlsx.writeBuffer() as Buffer;
    const id = newExportToken();
    preparedExports.set(id, {
      buffer,
      filename: `AT242_Rapport_Capabilite_${sanitizeFilename(productRef) || "PRODUIT"}.xlsx`,
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      createdAt: Date.now()
    });
    
    // Automatically delete from memory after 10 minutes
    setTimeout(() => {
      preparedExports.delete(id);
    }, 10 * 60 * 1000);
    
    res.json({ success: true, id });
  } catch (err: any) {
    console.error("Error generating capability Excel:", err);
    res.status(500).json({ error: "Failed to generate Excel report: " + err.message });
  }
});

// API Endpoint to export Repeatability (GRR) to Excel using exceljs
app.post("/api/export/repeatability", requireAuth, async (req, res) => {
  try {
    const { results, rawMeasurements, productRef, testerName, operatorName, files, testName } = req.body;
    
    const candidates = [
      path.resolve(process.cwd(), 'repet_report.xlsx'),
      path.join(DATA_DIR, 'repet_report.xlsx')
    ];
    const templatePath = candidates.find(p => {
      try {
        return fs.statSync(p).isFile();
      } catch {
        return false;
      }
    });
    
    const workbook = new ExcelJS.Workbook();
    if (templatePath) {
      await workbook.xlsx.readFile(templatePath);
    } else {
      console.warn(`repet_report.xlsx not found at expected paths. Creating fallback workbook dynamically.`);
      workbook.addWorksheet("Header");
      workbook.addWorksheet("Raw Data");
    }
    
    const wsHeader = workbook.getWorksheet("Header");
    const wsRaw = workbook.getWorksheet("Raw Data");
    
    if (wsHeader) {
      wsHeader.getCell("D10").value = testerName + " - " + productRef;
      wsHeader.getCell("D11").value = productRef;
      wsHeader.getCell("D12").value = "V1.0";
      wsHeader.getCell("C17").value = testerName;
      wsHeader.getCell("E17").value = testName || (testerName + " Bench");
      wsHeader.getCell("D39").value = operatorName;
      wsHeader.getCell("F39").value = new Date().toLocaleDateString('fr-FR');
      
      const allOk = results.every((r: any) => r.status === "OK");
      wsHeader.getCell("C21").value = allOk ? "OK" : "NOK";
    }
    
    if (wsRaw) {
      for (let r = 13; r <= 200; r++) {
        const row = wsRaw.getRow(r);
        row.eachCell((cell) => {
          cell.value = null;
        });
      }
      
      let assignments: any[] = [];
      if (files && files.length > 0) {
        const sortedFiles = [...files].sort((a: any, b: any) => a.filename.localeCompare(b.filename));
        assignments = sortedFiles.map((f: any, idx: number) => assignGrrMetadataBackend(f.filename, idx, sortedFiles.length));
      }
      
      for (let idx = 0; idx < results.length; idx++) {
        const r = results[idx];
        const key = `${r.testPointId}::${r.testStepName || ""}`;
        const vals = rawMeasurements[key] || [];
        const rowIndex = 13 + idx;
        const row = wsRaw.getRow(rowIndex);
        
        row.getCell(2).value = r.testPointId; // Col B
        row.getCell(3).value = r.testStepName; // Col C
        row.getCell(4).value = r.lsl; // Col D
        row.getCell(5).value = r.usl; // Col E
        row.getCell(6).value = r.unit || ""; // Col F
        
        if (assignments.length > 0) {
          const matrix: number[][][] = [
            [[], []], // Operator 0: Part 0, Part 1
            [[], []]  // Operator 1: Part 0, Part 1
          ];
          for (let i = 0; i < assignments.length; i++) {
            const assign = assignments[i];
            const val = vals[i];
            if (val !== undefined && assign) {
              matrix[assign.operatorIndex][assign.partIndex].push(val);
            }
          }
          
          // Op 1, Prod 1 -> columns G, H, I (columns 7, 8, 9)
          const op1_prod1 = matrix[0][0];
          op1_prod1.forEach((val, valIdx) => {
            row.getCell(7 + valIdx).value = val;
          });
          
          // Op 2, Prod 1 -> columns J, K, L (columns 10, 11, 12)
          const op2_prod1 = matrix[1][0];
          op2_prod1.forEach((val, valIdx) => {
            row.getCell(10 + valIdx).value = val;
          });
          
          // Op 1, Prod 2 -> columns M, N, O (columns 13, 14, 15)
          const op1_prod2 = matrix[0][1];
          op1_prod2.forEach((val, valIdx) => {
            row.getCell(13 + valIdx).value = val;
          });
          
          // Op 2, Prod 2 -> columns P, Q, R (columns 16, 17, 18)
          const op2_prod2 = matrix[1][1];
          op2_prod2.forEach((val, valIdx) => {
            row.getCell(16 + valIdx).value = val;
          });
        } else {
          for (let valIdx = 0; valIdx < vals.length; valIdx++) {
            row.getCell(7 + valIdx).value = vals[valIdx];
          }
        }
        row.commit();
      }
    }
    
    const buffer = await workbook.xlsx.writeBuffer() as Buffer;
    const id = newExportToken();
    preparedExports.set(id, {
      buffer,
      filename: `AT243_Rapport_Repetabilite_${sanitizeFilename(productRef) || "PRODUIT"}.xlsx`,
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      createdAt: Date.now()
    });
    
    // Automatically delete from memory after 10 minutes
    setTimeout(() => {
      preparedExports.delete(id);
    }, 10 * 60 * 1000);
    
    res.json({ success: true, id });
  } catch (err: any) {
    console.error("EXCEL EXPORT ERROR:", err);
    res.status(500).json({ error: "L'exportation a échoué côté serveur.", details: err.message });
  }
});

// API Endpoint to prepare generic exports from client-side (e.g. Airbus report)
app.post("/api/export/prepare-generic", requireAuth, (req, res) => {
  try {
    const { base64Data, filename, contentType } = req.body;
    if (!base64Data) {
      return res.status(400).json({ error: "Missing base64Data" });
    }
    const buffer = Buffer.from(base64Data, 'base64');
    const id = newExportToken();
    preparedExports.set(id, {
      buffer,
      filename: sanitizeFilename(filename) || "export.xlsx",
      contentType: contentType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      createdAt: Date.now()
    });
    
    // Automatically delete from memory after 10 minutes
    setTimeout(() => {
      preparedExports.delete(id);
    }, 10 * 60 * 1000);
    
    res.json({ success: true, id });
  } catch (err: any) {
    console.error("Error preparing generic export:", err);
    res.status(500).json({ error: err.message });
  }
});

// GET Endpoint to download prepared exports securely bypassing iframe sandboxing limits
app.get("/api/export/download/:id", (req, res) => {
  const id = req.params.id;
  // Strict format gate: only 64-hex crypto tokens are valid
  if (!/^[a-f0-9]{64}$/.test(id)) {
    return res.status(400).send("Token invalide.");
  }
  const prepared = preparedExports.get(id);
  if (!prepared) {
    return res.status(404).send("Ce rapport a expiré ou n'existe pas. Veuillez recalculer et réessayer.");
  }
  // Enforce 10-minute freshness
  if (Date.now() - prepared.createdAt > 10 * 60 * 1000) {
    preparedExports.delete(id);
    return res.status(410).send("Ce rapport a expiré. Veuillez recalculer et réessayer.");
  }

  res.setHeader("Content-Type", prepared.contentType);
  res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(prepared.filename)}`);
  res.send(prepared.buffer);

  // Clear memory immediately after download
  preparedExports.delete(id);
});

// API Endpoint to get masters list loaded from Excel
app.get("/api/masters", requireAuth, (req, res) => {
  const masters = loadMastersFromExcel();
  res.json(masters);
});

// API Endpoint to update and write masters back to Excel (any authenticated user)
app.post("/api/masters", requireAuth, (req, res) => {
  const { masters } = req.body;
  if (!Array.isArray(masters)) {
    return res.status(400).json({ success: false, error: "Invalid masters data. Must be an array." });
  }
  const success = saveMastersToExcel(masters);
  if (success) {
    res.json({ success: true, message: "Masters successfully saved to Excel." });
  } else {
    res.status(500).json({ success: false, error: "Failed to write masters to Excel file." });
  }
});

// API Endpoint to create a single new master and persist it to the Excel file.
// It loads the current masters from the physical Excel file, appends the new
// one, and writes the updated list straight back so the file stays in sync.
app.post("/api/masters/new", requireAuth, (req, res) => {
  try {
    const incoming = req.body;
    if (!incoming || typeof incoming !== "object" || !incoming.idMaster) {
      return res.status(400).json({ success: false, error: "Invalid master data. 'idMaster' is required." });
    }

    const masters = loadMastersFromExcel();
    const newMaster = {
      id: String(incoming.idMaster),
      testeur: incoming.testeur || "",
      idMaster: incoming.idMaster,
      refProduitMaster: incoming.refProduitMaster || "",
      numSerieProduitMaster: incoming.numSerieProduitMaster || "",
      refCarteMaster: incoming.refCarteMaster || "",
      numSerieCarteMaster: incoming.numSerieCarteMaster || "",
      dateCreation: incoming.dateCreation || new Date().toLocaleDateString('fr-FR'),
      commentaire1: incoming.commentaire1 || "",
      statut: incoming.statut || "Active",
      commentaire2: incoming.commentaire2 || "",
      verif: incoming.verif || "OK"
    };

    masters.push(newMaster);
    const success = saveMastersToExcel(masters);

    if (success) {
      // Resync the data source from the written Excel file so the UI updates instantly.
      const updated = resyncMastersToDataSource();
      res.json({ success: true, message: "New master created and saved to Excel.", master: newMaster, masters: updated });
    } else {
      res.status(500).json({ success: false, error: "Failed to write new master to Excel file." });
    }
  } catch (error: any) {
    console.error("SERVER_SAVE_ERROR:", error);
    res.status(500).json({ error: error.message });
  }
});

// ---------------------------------------------------------------------------
// Authentication, session & user-management routes (SQLite-backed).
// Passwords are hashed with bcryptjs (per-user random salt) and stored ONLY as
// `password_hash` in the `users` table (see src/db.ts). Sessions are stateless
// JWTs (HS256) carrying { role, matricule, name, jti, exp }; logout revokes the
// token's `jti` in the SQLite `revoked_tokens` table so it cannot be reused.
//
// The old file-based auth (technicians.json + ADMIN_CODE_HASH env secret) has
// been removed: the `users` table is now the single source of truth.
// ---------------------------------------------------------------------------

// Default seed credentials, overridable via env for first-boot flexibility.
// They are ONLY inserted when the `users` table is empty, and must be rotated
// after the first login (a warning is logged at startup).
const DEFAULT_TECH_PASSWORD = process.env.SEED_TECH_PASSWORD || "Technicien2024!";
const DEFAULT_ADMIN_PASSWORD = process.env.SEED_ADMIN_PASSWORD || "Admin2024!";

// Seed default technicians + an admin on first startup so the system is usable
// immediately. Existing accounts are never overwritten (seeding is idempotent).
function seedUsersIfEmpty() {
  if (countUsers() > 0) return;
  const seed: Array<{ matricule: string; nom: string; role: "technician" | "admin" }> = [
    { matricule: "366", nom: "BEN MANSOUR Samir", role: "technician" },
    { matricule: "195", nom: "BEN GHODHBENE Hamza", role: "technician" },
    { matricule: "122", nom: "TECH", role: "technician" },
    { matricule: "ADM01", nom: "ADMINISTRATEUR METROLOGIE", role: "admin" },
  ];
  for (const u of seed) {
    createUser({ matricule: u.matricule, nom: u.nom, role: u.role, passwordHash: hashPassword(u.role === "admin" ? DEFAULT_ADMIN_PASSWORD : DEFAULT_TECH_PASSWORD) });
  }
  console.log(`[auth] Base initialisée: ${seed.length} utilisateurs créés (admin: ${DEFAULT_ADMIN_PASSWORD}, technicien: ${DEFAULT_TECH_PASSWORD}).`);
  console.warn("[auth] CHANGEZ les mots de passe par défaut après le premier lancement.");
}

// POST /api/auth/login  -- public (requireAuth exempts this route on every env).
// Body: { matricule, password, role? }  role defaults to 'technician'.
// Returns { token, user } on success.
app.post("/api/auth/login", requireAuth, async (req, res) => {
  try {
    const { matricule, password, role } = req.body || {};

    if (!matricule || typeof matricule !== "string") {
      return res.status(400).json({ error: "Le matricule est requis." });
    }
    if (!password || typeof password !== "string") {
      return res.status(400).json({ error: "Le mot de passe est requis." });
    }
    // A technician cannot authenticate as admin and vice-versa.
    const requestedRole: "technician" | "admin" = role === "admin" ? "admin" : "technician";

    const user = getUserByMatricule(matricule);
    // Single, non-revealing failure message to avoid user-enumeration.
    if (!user || user.role !== requestedRole || !comparePassword(password, user.password_hash)) {
      return res.status(401).json({ error: "Matricule ou mot de passe incorrect." });
    }

    const token = signToken({ role: user.role, matricule: user.matricule, name: user.nom });
    return res.json({ token, user: { role: user.role, matricule: user.matricule, name: user.nom } });
  } catch (error: any) {
    console.error("SERVER_LOGIN_ERROR:", error);
    return res.status(500).json({ error: "Erreur serveur lors de la connexion.", details: error?.message || "Erreur interne" });
  }
});

// POST /api/auth/logout -- invalidate the current session token (server-side).
app.post("/api/auth/logout", requireAuth, (req: any, res: any) => {
  try {
    const claims = req.auth;
    if (claims && claims.jti) {
      revokeToken(claims.jti, claims.exp || 0);
    }
    return res.json({ success: true, message: "Déconnexion réussie." });
  } catch (error: any) {
    console.error("SERVER_LOGOUT_ERROR:", error);
    return res.status(500).json({ error: "Erreur lors de la déconnexion.", details: error?.message || "Erreur interne" });
  }
});

// GET /api/auth/session -- return the authenticated user (or 401).
app.get("/api/auth/session", requireAuth, (req: any, res: any) => {
  try {
    const claims = req.auth;
    return res.json({ user: { role: claims.role, matricule: claims.matricule, name: claims.name } });
  } catch (error: any) {
    return res.status(500).json({ error: "Erreur serveur.", details: error?.message || "Erreur interne" });
  }
});

// Re-verify an administrator's password before a sensitive action (replaces the
// old ADMIN_CODE_HASH env-secret: the password now lives in the users table).
app.post("/api/auth/admin-verify", requireAuth, async (req: any, res: any) => {
  try {
    const claims = req.auth;
    if (claims.role !== "admin") {
      return res.status(403).json({ error: "Accès refusé: privilège administrateur requis." });
    }
    const { code } = req.body || {};
    if (typeof code !== "string" || code.length === 0 || code.length > 200) {
      return res.status(401).json({ error: "Code incorrect." });
    }
    const user = getUserByMatricule(claims.matricule);
    if (!user || !comparePassword(code, user.password_hash)) {
      return res.status(401).json({ error: "Code incorrect." });
    }
    return res.json({ success: true });
  } catch (error: any) {
    console.error("SERVER_ADMIN_VERIFY_ERROR:", error);
    return res.status(500).json({ error: "Erreur serveur.", details: error?.message || "Erreur interne" });
  }
});

// --- Technicians management (SQLite-backed) ---
// Public list of technicians is intentionally reachable without auth so the
// login screen can display available accounts. Only matricule + name are
// returned (never the password hash).
app.get("/api/technicians", (req, res) => {
  try {
    res.json(listTechnicians());
  } catch (error: any) {
    console.error("SERVER_TECHNICIANS_LIST_ERROR:", error);
    res.status(500).json({ error: error.message || "Erreur serveur" });
  }
});

// Create a technician. Admin-only (role-gated).
app.post("/api/technicians", requireRole("admin"), (req, res) => {
  try {
    const { matricule, name, password } = req.body || {};
    if (!matricule || !name || !password) {
      return res.status(400).json({ success: false, error: "Matricule, nom et mot de passe sont requis." });
    }
    if (getUserByMatricule(matricule)) {
      return res.status(409).json({ success: false, error: "Un utilisateur avec ce matricule existe déjà." });
    }
    createUser({ matricule, nom: name, role: "technician", passwordHash: hashPassword(password) });
    res.json({ success: true, technicians: listTechnicians() });
  } catch (error: any) {
    console.error("SERVER_TECHNICIAN_SAVE_ERROR:", error);
    res.status(500).json({ error: error.message || "Erreur serveur" });
  }
});

// Delete a technician. Admin-only (role-gated).
app.delete("/api/technicians/:matricule", requireRole("admin"), (req, res) => {
  try {
    const { matricule } = req.params;
    if (!matricule) {
      return res.status(400).json({ success: false, error: "Matricule requis." });
    }
    const ok = deleteTechnician(matricule);
    if (!ok) return res.status(404).json({ success: false, error: "Technicien introuvable." });
    res.json({ success: true, technicians: listTechnicians() });
  } catch (error: any) {
    console.error("SERVER_TECHNICIAN_DELETE_ERROR:", error);
    res.status(500).json({ error: error.message || "Erreur serveur" });
  }
});

// Vite dev server middleware in non-production, otherwise static build serving
async function startServer() {
  console.log("Initializing database/Excel sync to /src/data.ts on startup...");
  loadMastersFromExcel();
  // Initialize the SQLite user store & seed default accounts on first boot.
  purgeExpiredRevocations();
  seedUsersIfEmpty();

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
    console.log("Vite dev middleware mounted successfully.");
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
    console.log("Production static server mounted for: " + distPath);
  }

  // Bind to loopback by default; only expose when SERVER_HOST is explicitly set
 // const HOST = process.env.SERVER_HOST || "127.0.0.1";
  //app.listen(PORT, HOST, () => {
    //console.log(`Server running on ${HOST}:${PORT}`);
  //});
  const PORT = process.env.PORT || 3005;

// تأكد من إضافة '0.0.0.0' لتسمح لـ Render بالاتصال بالسيرفر
app.listen(Number(PORT), '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
});
}

startServer();
