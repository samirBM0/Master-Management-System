import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";
import { 
  Search, 
  Plus, 
  Trash2, 
  Edit3, 
  Calculator, 
  TrendingUp, 
  CheckCircle, 
  AlertTriangle, 
  XCircle, 
  PlusCircle, 
  Download, 
  FileSpreadsheet, 
  User, 
  Calendar, 
  Copy, 
  Check, 
  RefreshCw, 
  FileText, 
  Info, 
  Gauge, 
  Layers, 
  Filter,
  ArrowRight,
  Sparkles,
  Settings,
  Upload,
  Cpu,
  Terminal,
  ExternalLink,
  Link2,
  Play,
  X,
  ChevronDown,
  Shield,
  Lock,
  Unlock,
  LogIn,
  LogOut,
  Key
} from "lucide-react";
import { INITIAL_MASTER_DATA } from "./data";
import { MasterItem, CapabilityCalculation, UserSession, HistoryLog } from "./types";
import ActipaPanel from "./actipa/ActipaPanel";
import RadarChart from "./components/RadarChart";

// Read the session token from storage (opaque, server-signed). The API_TOKEN
// client constant is unused now that the server issues per-session tokens.
const API_TOKEN: string | undefined = (import.meta as any).env?.VITE_API_TOKEN;

// Build the Authorization header (Bearer token) if a token is configured.
function authHeaders(): Record<string, string> {
  const h: Record<string, string> = { "Content-Type": "application/json" };
  const tok = (typeof localStorage !== "undefined") ? localStorage.getItem("mms_token") : null;
  const bearer = tok || API_TOKEN;
  if (bearer) h["Authorization"] = `Bearer ${bearer}`;
  return h;
}

// Decode a V9 opaque signed session token payload (base64url JWT-like)
function decodeV9Token(token: string): any | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return null;
    const base64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const padded = base64 + "=".repeat((4 - (base64.length % 4)) % 4);
    const json = decodeURIComponent(
      atob(padded)
        .split("")
        .map(c => "%" + c.charCodeAt(0).toString(16).padStart(2, "0"))
        .join("")
    );
    return JSON.parse(json);
  } catch {
    return null;
  }
}

// Defense-in-depth against SheetJS prototype-pollution (CVE-2023-30533):
// neutralize dangerous __proto__ / constructor.prototype keys before/after
// parsing untrusted workbooks. SheetJS npm latest is 0.18.5; patched builds
// ship via the SheetJS CDN, so we harden parsing on the client instead.
function sanitizeProto(obj: any): any {
  if (Array.isArray(obj)) return obj.map(sanitizeProto);
  if (obj && typeof obj === "object") {
    const out: any = {};
    for (const k of Object.keys(obj)) {
      if (k === "__proto__" || k === "constructor" || k === "prototype") continue;
      out[k] = sanitizeProto(obj[k]);
    }
    return out;
  }
  return obj;
}
function safeReadXlsx(data: any, opts: any): any {
  const wb = XLSX.read(data, opts);
  if (wb && Array.isArray(wb.SheetNames)) {
    wb.SheetNames = wb.SheetNames.map((n: string) => String(n));
  }
  return wb;
}

const formatExcelDate = (val: any): string => {
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
};

const ActiaLogo = ({ className = "h-10" }: { className?: string }) => (
  <svg 
    viewBox="0 0 240 80" 
    className={className} 
    fill="none" 
    xmlns="http://www.w3.org/2000/svg"
  >
    {/* Background slanted block for ACTIA (parallelogram) */}
    <path 
      d="M 12 10 L 175 10 L 160 62 L -3 62 Z" 
      fill="#5C5959" 
    />
    
    {/* ACTIA Text slanted */}
    <text 
      x="20" 
      y="48" 
      fill="#FFFFFF" 
      fontSize="36" 
      fontWeight="900" 
      fontFamily="system-ui, -apple-system, sans-serif" 
      fontStyle="italic"
      letterSpacing="-1"
    >
      ACTIA
    </text>
    
    {/* Registered Trademark symbol ® */}
    <text 
      x="142" 
      y="28" 
      fill="#FFFFFF" 
      fontSize="10" 
      fontWeight="bold" 
      fontFamily="system-ui, -apple-system, sans-serif"
    >
      ®
    </text>
    
    {/* Slanted Green Grid (parallelogram shape) on the right */}
    <g transform="translate(142, 0)">
      {/* Rhombuses in Row 1 */}
      <path d="M 38 10 L 51 10 L 46 24 L 33 24 Z" fill="#00965E" />
      <path d="M 53 10 L 66 10 L 61 24 L 48 24 Z" fill="#00965E" />
      <path d="M 68 10 L 81 10 L 76 24 L 63 24 Z" fill="#00965E" />
      <path d="M 83 10 L 96 10 L 91 24 L 78 24 Z" fill="#00965E" />

      {/* Rhombuses in Row 2 */}
      <path d="M 33 27 L 46 27 L 41 41 L 28 41 Z" fill="#00965E" />
      <path d="M 48 27 L 61 27 L 56 41 L 43 41 Z" fill="#00965E" />
      <path d="M 63 27 L 76 27 L 71 41 L 58 41 Z" fill="#00965E" />
      <path d="M 78 27 L 91 27 L 86 41 L 73 41 Z" fill="#00965E" />

      {/* Rhombuses in Row 3 */}
      <path d="M 28 44 L 41 44 L 36 58 L 23 58 Z" fill="#00965E" />
      <path d="M 43 44 L 56 44 L 51 58 L 38 58 Z" fill="#00965E" />
      <path d="M 58 44 L 71 44 L 66 58 L 53 58 Z" fill="#00965E" />
      <path d="M 73 44 L 86 44 L 81 58 L 68 58 Z" fill="#00965E" />
    </g>

    {/* TUNISIE text at the bottom */}
    <text 
      x="100" 
      y="75" 
      fill="#5C5959" 
      fontSize="16" 
      fontWeight="700" 
      fontFamily="system-ui, -apple-system, sans-serif" 
      fontStyle="italic"
      letterSpacing="4"
    >
      TUNISIE
    </text>
  </svg>
);

const INITIAL_CALCULATIONS: CapabilityCalculation[] = [
  // MASTER_175: Drifting measurements
  {
    id: "CALC_175_1",
    title: "Vérification Capabilité Fréquence RF",
    date: "12/07/2026",
    testeur: "UFT SMOOVE RTS",
    masterId: "MASTER_175",
    characteristic: "Fréquence RF (MHz)",
    nominal: 433.42,
    lsl: 433.22,
    usl: 433.62,
    measurements: [433.41, 433.43, 433.40, 433.42, 433.41, 433.43, 433.42, 433.39, 433.41, 433.42, 433.43, 433.40, 433.42, 433.41, 433.43],
    operator: "BEN MANSOUR Samir"
  },
  {
    id: "CALC_175_2",
    title: "Vérification Capabilité Fréquence RF",
    date: "13/07/2026",
    testeur: "UFT SMOOVE RTS",
    masterId: "MASTER_175",
    characteristic: "Fréquence RF (MHz)",
    nominal: 433.42,
    lsl: 433.22,
    usl: 433.62,
    measurements: [433.43, 433.45, 433.42, 433.44, 433.43, 433.45, 433.44, 433.41, 433.43, 433.44, 433.45, 433.42, 433.44, 433.43, 433.45],
    operator: "BEN MANSOUR Samir"
  },
  {
    id: "CALC_175_3",
    title: "Vérification Capabilité Fréquence RF",
    date: "14/07/2026",
    testeur: "UFT SMOOVE RTS",
    masterId: "MASTER_175",
    characteristic: "Fréquence RF (MHz)",
    nominal: 433.42,
    lsl: 433.22,
    usl: 433.62,
    measurements: [433.46, 433.48, 433.45, 433.47, 433.46, 433.48, 433.47, 433.44, 433.46, 433.47, 433.48, 433.45, 433.47, 433.46, 433.48],
    operator: "BEN MANSOUR Samir"
  },
  // MASTER_176: Stable measurements (Tension Régulée)
  {
    id: "CALC_176_1",
    title: "Contrôle Tension Régulée",
    date: "12/07/2026",
    testeur: "UFT SMOOVE RTS",
    masterId: "MASTER_176",
    characteristic: "Tension Régulée (V)",
    nominal: 5.00,
    lsl: 4.75,
    usl: 5.25,
    measurements: [5.00, 4.99, 5.01, 5.00, 4.98, 5.02, 5.00, 4.99, 5.01, 5.00, 4.99, 5.01, 5.00, 4.98, 5.02],
    operator: "DURAND Nicolas"
  },
  {
    id: "CALC_176_2",
    title: "Contrôle Tension Régulée",
    date: "13/07/2026",
    testeur: "UFT SMOOVE RTS",
    masterId: "MASTER_176",
    characteristic: "Tension Régulée (V)",
    nominal: 5.00,
    lsl: 4.75,
    usl: 5.25,
    measurements: [5.01, 5.00, 5.02, 5.01, 4.99, 5.03, 5.01, 5.00, 5.02, 5.01, 5.00, 5.02, 5.01, 4.99, 5.03],
    operator: "DURAND Nicolas"
  },
  {
    id: "CALC_176_3",
    title: "Contrôle Tension Régulée",
    date: "14/07/2026",
    testeur: "UFT SMOOVE RTS",
    masterId: "MASTER_176",
    characteristic: "Tension Régulée (V)",
    nominal: 5.00,
    lsl: 4.75,
    usl: 5.25,
    measurements: [4.99, 4.98, 5.00, 4.99, 4.97, 5.01, 4.99, 4.98, 5.00, 4.99, 4.98, 5.00, 4.99, 4.97, 5.01],
    operator: "DURAND Nicolas"
  }
];

export default function App() {
  // Tabs: 'testeurs' | 'nouveau-master' | 'calcul' | 'admin-panel'
  const [activeTab, setActiveTab] = useState<'testeurs' | 'nouveau-master' | 'calcul' | 'admin-panel'>('testeurs');
  
  // PDF export loading state
  const [isExportingPDF, setIsExportingPDF] = useState(false);

  const [mastersLoaded, setMastersLoaded] = useState(false);

  const [masters, setMasters] = useState<MasterItem[]>(INITIAL_MASTER_DATA);

  // Fetch from server on mount
  useEffect(() => {
    fetch("/api/masters", { headers: authHeaders() })
      .then(res => {
        if (!res.ok) throw new Error("Server response was not ok");
        return res.json();
      })
      .then(data => {
        if (data && Array.isArray(data)) {
          setMasters(data);
        }
        setMastersLoaded(true);
      })
      .catch(err => {
        console.error("Failed to load masters from server, using local data", err);
        setMastersLoaded(true); // Fallback to initial
      });
  }, []);

  // Save masters whenever state changes
  useEffect(() => {
    localStorage.setItem("master_management_items", JSON.stringify(masters));

    if (!mastersLoaded) return;

    fetch("/api/masters", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({ masters })
    })
    .then(res => {
      if (!res.ok) throw new Error("Server write failed");
      return res.json();
    })
    .then(data => {
      if (!data.success) {
        console.error("Failed to save masters on server:", data.error);
      } else {
        console.log("Masters synced to server successfully");
      }
    })
    .catch(err => {
      console.error("Network error saving masters on server", err);
    });
  }, [masters, mastersLoaded]);

  // Capability calculations history
  const [calculations, setCalculations] = useState<CapabilityCalculation[]>(() => {
    const saved = localStorage.getItem("mms_capability_calculations");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error("Error reading saved calculations", e);
      }
    }
    return INITIAL_CALCULATIONS;
  });

  useEffect(() => {
    localStorage.setItem("mms_capability_calculations", JSON.stringify(calculations));
  }, [calculations]);

  // Compute master drift trend based on 3 latest calculations
  const getMasterTrend = (masterId: string) => {
    const masterCalcs = calculations.filter(c => c.masterId === masterId);
    if (masterCalcs.length < 3) return null;

    // Take the 3 latest calculations
    const latestCalcs = masterCalcs.slice(-3);

    // Compute mean for each of the 3 latest calculations
    const means = latestCalcs.map(calc => {
      const sum = calc.measurements.reduce((acc, v) => acc + v, 0);
      return sum / calc.measurements.length;
    });

    const averageMean = means.reduce((acc, m) => acc + m, 0) / 3;

    // Use parameters from the most recent calculation
    const latestCalc = latestCalcs[2];
    const { nominal, lsl, usl } = latestCalc;
    const tolerance = usl - lsl;

    if (tolerance <= 0) return null;

    // Drift percentage of the average mean relative to the tolerance interval
    const drift = Math.abs(averageMean - nominal);
    const driftPercentage = (drift / tolerance) * 100;

    const isWarning = driftPercentage > 5.0;

    return {
      isWarning,
      driftPercentage: parseFloat(driftPercentage.toFixed(2)),
      averageMean: parseFloat(averageMean.toFixed(4)),
      nominal,
      lsl,
      usl,
      tolerance,
      means: means.map(m => parseFloat(m.toFixed(4))),
      latestDate: latestCalc.date
    };
  };

  // Drifting masters count
  const driftingMastersCount = useMemo(() => {
    return masters.filter(m => {
      const trend = getMasterTrend(m.idMaster);
      return trend && trend.isWarning;
    }).length;
  }, [masters, calculations]);


  // List of unique testers for easy dropdowns/filtering
  const uniqueTesters = useMemo(() => {
    const set = new Set(masters.map(m => m.testeur?.trim()).filter(Boolean));
    return Array.from(set).sort();
  }, [masters]);

  // Toast notification state
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'info' | 'error' } | null>(null);
  const showToast = (message: string, type: 'success' | 'info' | 'error' = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  };

  // State for search and filter in "1. Testeur" (Applied Filters used for computed lists)
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedTesterFilter, setSelectedTesterFilter] = useState<string>("All");
  const [statusFilter, setStatusFilter] = useState<string>("All");
  const [verifFilter, setVerifFilter] = useState<string>("All");

  // Temp/Draft state for interactive inputs before clicking validation button
  const [searchQueryInput, setSearchQueryInput] = useState("");
  const [selectedTesterFilterInput, setSelectedTesterFilterInput] = useState<string>("All");
  const [statusFilterInput, setStatusFilterInput] = useState<string>("All");
  const [verifFilterInput, setVerifFilterInput] = useState<string>("All");

  // Apply draft filters to active filter state
  const handleApplyFilters = () => {
    setSearchQuery(searchQueryInput);
    setSelectedTesterFilter(selectedTesterFilterInput);
    setStatusFilter(statusFilterInput);
    setVerifFilter(verifFilterInput);
    showToast("Filtres de recherche appliqués avec succès !", "success");
  };

  // Helper to determine if there are any unapplied changes in the filter inputs
  const hasPendingFilterChanges = useMemo(() => {
    return searchQuery !== searchQueryInput ||
      selectedTesterFilter !== selectedTesterFilterInput ||
      statusFilter !== statusFilterInput ||
      verifFilter !== verifFilterInput;
  }, [
    searchQuery, searchQueryInput,
    selectedTesterFilter, selectedTesterFilterInput,
    statusFilter, statusFilterInput,
    verifFilter, verifFilterInput
  ]);

  // State for creating a new master
  const getNextMasterId = () => {
    const ids = masters
      .map(m => m.idMaster)
      .filter(id => id.startsWith("MASTER_"))
      .map(id => parseInt(id.replace("MASTER_", ""), 10))
      .filter(num => !isNaN(num));
    
    const max = ids.length > 0 ? Math.max(...ids) : 451;
    const nextNum = max + 1;
    return `MASTER_${String(nextNum).padStart(3, '0')}`;
  };

  const [newMaster, setNewMaster] = useState<Partial<MasterItem>>({
    testeur: "",
    idMaster: "",
    refProduitMaster: "",
    numSerieProduitMaster: "",
    refCarteMaster: "",
    numSerieCarteMaster: "",
    dateCreation: new Date().toLocaleDateString('fr-FR'),
    commentaire1: "",
    statut: "Active",
    commentaire2: "",
    verif: "OK"
  });

  // Automatically compute next available ID when we navigate to create or list changes
  useEffect(() => {
    if (activeTab === 'nouveau-master' && !newMaster.idMaster) {
      setNewMaster(prev => ({ ...prev, idMaster: getNextMasterId() }));
    }
  }, [activeTab, masters]);

  // Raw CSV Paste State for advanced batch import
  const [rawCsvPaste, setRawCsvPaste] = useState("");
  const [showCsvImporter, setShowCsvImporter] = useState(false);

  // Helper to normalize status strings for robust, accent-insensitive and gender-insensitive comparisons
  const normalizeStatus = (s: string): string => {
    if (!s) return "";
    let clean = s.toLowerCase().trim();
    
    // Replace typical corrupted characters safely without empty regex literals
    clean = clean.replace(new RegExp("[]", "g"), "e");
    
    // Normalize standard accents
    clean = clean.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    if (clean.includes("active") || clean.includes("actif")) return "active";
    if (clean.includes("obsol")) return "obsolete";
    if (clean.includes("endom")) return "endommagee";
    return clean;
  };

  // State for Excel/CSV Database Importer (to keep the list of masters up-to-date)
  const [isDraggingExcel, setIsDraggingExcel] = useState(false);
  const [excelImportDetails, setExcelImportDetails] = useState<{
    fileName: string;
    rowCount: number;
    parsedItems: MasterItem[];
  } | null>(null);

  // Parse raw text of a CSV/Excel sheet containing semicolon-separated values
  const parseUploadedCsvContent = (text: string): MasterItem[] => {
    try {
      const workbook = safeReadXlsx(text, { type: "string" });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

      // Find header row containing TESTEURS
      let headerIdx = -1;
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        if (row && row[0] && String(row[0]).trim().toUpperCase() === "TESTEURS") {
          headerIdx = i;
          break;
        }
      }

      const parsed: MasterItem[] = [];
      const startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;

      for (let i = startIdx; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        // Skip metadata/header lines if we didn't find headerIdx but they look like it
        if (row[0] && (
          String(row[0]).toLowerCase().includes("suivi des") ||
          String(row[0]).toLowerCase().includes("m-à-j") ||
          String(row[0]).toLowerCase().includes("m-a-j") ||
          String(row[0]).toLowerCase().includes("date:") ||
          String(row[0]).toLowerCase().includes("verion:") ||
          String(row[0]).toLowerCase().includes("version:")
        )) {
          continue;
        }

        const idMaster = String(row[1] || "").trim();
        if (!idMaster || idMaster.toLowerCase() === "identificationmaster") continue; // Skip empty/header row

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
      return parsed;
    } catch (e) {
      console.error("Error parsing CSV content:", e);
      return [];
    }
  };

  const handleExcelFileParse = (file: File) => {
    const reader = new FileReader();
    const isXlsx = file.name.endsWith(".xlsx") || file.name.endsWith(".xls");
    
    reader.onload = (e) => {
      try {
        let items: MasterItem[] = [];
        if (isXlsx) {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = safeReadXlsx(data, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
          
          // Find header row containing TESTEURS
          let headerIdx = -1;
          for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (row && row[0] && String(row[0]).trim().toUpperCase() === "TESTEURS") {
              headerIdx = i;
              break;
            }
          }

          const startIdx = headerIdx !== -1 ? headerIdx + 1 : 0;
          for (let i = startIdx; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;

            const idMaster = String(row[1] || "").trim();
            if (!idMaster || idMaster.toLowerCase() === "identificationmaster") continue;

            items.push({
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
        } else {
          const textContent = e.target?.result as string;
          items = parseUploadedCsvContent(textContent);
        }

        if (items.length > 0) {
          setExcelImportDetails({
            fileName: file.name,
            rowCount: items.length,
            parsedItems: items
          });
          showToast(`${items.length} pièces étalons trouvées. Prêt à importer !`, "success");
        } else {
          showToast("Aucune ligne de données valide détectée. Vérifiez le format du fichier.", "error");
        }
      } catch (err) {
        console.error(err);
        showToast("Erreur lors du traitement du fichier Excel.", "error");
      }
    };

    if (isXlsx) {
      reader.readAsArrayBuffer(file);
    } else {
      reader.readAsText(file);
    }
  };

  const handleApplyDatabaseImport = (mode: 'replace' | 'merge') => {
    if (!excelImportDetails) return;
    
    if (mode === 'replace') {
      setMasters(excelImportDetails.parsedItems);
      showToast(`Base de données réinitialisée ! ${excelImportDetails.parsedItems.length} pièces chargées.`, "success");
    } else {
      setMasters(prev => {
        // Create a map by idMaster to quickly match and override or insert
        const map = new Map(prev.map(m => [m.idMaster.toUpperCase(), m]));
        excelImportDetails.parsedItems.forEach(item => {
          map.set(item.idMaster.toUpperCase(), item);
        });
        return Array.from(map.values());
      });
      showToast(`${excelImportDetails.parsedItems.length} pièces synchronisées dans la base existante.`, "success");
    }
    setExcelImportDetails(null);
  };

  // Filtered masters list based on applied filters
  const filteredMasters = useMemo(() => {
    return masters.filter(item => {
      const matchSearch = !searchQuery || (
        (item.idMaster && item.idMaster.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (item.testeur && item.testeur.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (item.refProduitMaster && item.refProduitMaster.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (item.numSerieProduitMaster && item.numSerieProduitMaster.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (item.refCarteMaster && item.refCarteMaster.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (item.commentaire1 && item.commentaire1.toLowerCase().includes(searchQuery.toLowerCase())) ||
        (item.commentaire2 && item.commentaire2.toLowerCase().includes(searchQuery.toLowerCase()))
      );

      const matchTester = selectedTesterFilter === "All" || (item.testeur && item.testeur.toLowerCase().trim() === selectedTesterFilter.toLowerCase().trim());
      
      const matchStatus = 
        statusFilter === "All" || 
        normalizeStatus(item.statut) === normalizeStatus(statusFilter);

      const matchVerif = verifFilter === "All" || item.verif === verifFilter;

      return matchSearch && matchTester && matchStatus && matchVerif;
    });
  }, [masters, searchQuery, selectedTesterFilter, statusFilter, verifFilter]);

  // Summary counts using robust status normalization
  const stats = useMemo(() => {
    const total = filteredMasters.length;
    const active = filteredMasters.filter(m => normalizeStatus(m.statut) === "active").length;
    const obsolete = filteredMasters.filter(m => normalizeStatus(m.statut) === "obsolete").length;
    const endommagee = filteredMasters.filter(m => normalizeStatus(m.statut) === "endommagee").length;
    const okVerif = filteredMasters.filter(m => m.verif === "OK").length;
    const okPercent = total > 0 ? Math.round((okVerif / total) * 100) : 0;

    return { total, active, obsolete, endommagee, okVerif, okPercent };
  }, [filteredMasters]);

  // Export menu state + handlers (CSV / XLSX)
  const [showExportMenu, setShowExportMenu] = useState(false);

  const exportMastersCSV = (items: MasterItem[]) => {
    const headers = "TESTEURS;Identification Master;Ref Produit;Num Serie;Ref Carte;Num Serie Carte;Date Creation;Commentaire;Statut;Verif\n";
    const rows = items.map(m =>
      `"${m.testeur}";"${m.idMaster}";"${m.refProduitMaster}";"${m.numSerieProduitMaster}";"${m.refCarteMaster}";"${m.numSerieCarteMaster}";"${m.dateCreation}";"${m.commentaire1}";"${m.statut}";"${m.verif}"`
    ).join("\n");
    const blob = new Blob([headers + rows], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `export_masters_${new Date().toISOString().split('T')[0]}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast("Fichier CSV exporté avec succès !", "success");
  };

  const exportMastersXLSX = (items: MasterItem[]) => {
    const data = items.map(m => ({
      "Testeur": m.testeur,
      "Identification Master": m.idMaster,
      "Référence Produit": m.refProduitMaster,
      "Num Série Produit": m.numSerieProduitMaster,
      "Référence Carte": m.refCarteMaster,
      "Num Série Carte": m.numSerieCarteMaster,
      "Date Création": m.dateCreation,
      "Commentaire": m.commentaire1,
      "Statut": m.statut,
      "Vérif": m.verif
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Masters");
    XLSX.writeFile(workbook, `export_masters_${new Date().toISOString().split('T')[0]}.xlsx`);
    showToast("Fichier Excel (.xlsx) exporté avec succès !", "success");
  };

  // Handle creating master
  const handleCreateMaster = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newMaster.testeur) {
      showToast("Veuillez sélectionner ou saisir un testeur", "error");
      return;
    }
    if (!newMaster.idMaster) {
      showToast("Veuillez saisir une identification Master", "error");
      return;
    }

    // Check duplicate
    if (masters.some(m => m.idMaster.toUpperCase() === newMaster.idMaster?.toUpperCase())) {
      showToast(`L'identifiant ${newMaster.idMaster} existe déjà !`, "error");
      return;
    }

    const created: MasterItem = {
      id: newMaster.idMaster,
      testeur: newMaster.testeur.trim(),
      idMaster: newMaster.idMaster.trim(),
      refProduitMaster: (newMaster.refProduitMaster || "").trim(),
      numSerieProduitMaster: (newMaster.numSerieProduitMaster || "").trim(),
      refCarteMaster: (newMaster.refCarteMaster || "").trim(),
      numSerieCarteMaster: (newMaster.numSerieCarteMaster || "").trim(),
      dateCreation: newMaster.dateCreation || new Date().toLocaleDateString('fr-FR'),
      commentaire1: (newMaster.commentaire1 || "").trim(),
      statut: newMaster.statut || "Active",
      commentaire2: (newMaster.commentaire2 || "").trim(),
      verif: newMaster.verif || "OK"
    };

    setMasters(prev => [created, ...prev]);
    showToast(`Master ${created.idMaster} enregistré avec succès !`, "success");
    addLog("Création Master", `Fiche Master créée - ID: ${created.idMaster}, Testeur: ${created.testeur}, Statut: ${created.statut}`);

    // Explicitly persist the new master to the backend so it is written back
    // to the physical Excel file (POST /api/masters/new -> saveMastersToExcel).
    // Include the V9 session token explicitly, mirroring the login session logic.
    const sessionToken = session?.token || localStorage.getItem("mms_token") || API_TOKEN;
    fetch("/api/masters/new", {
      method: "POST",
      headers: sessionToken
        ? { "Content-Type": "application/json", Authorization: `Bearer ${sessionToken}` }
        : authHeaders(),
      body: JSON.stringify(created)
    })
      .then(res => res.ok ? res.json() : Promise.reject(new Error("Server write failed")))
      .then(data => {
        if (!data.success) {
          console.error("Failed to save new master on server:", data.error);
          showToast("Master enregistré localement mais non sauvegardé sur le serveur", "error");
        } else {
          console.log("New master saved to Excel successfully");
        }
      })
      .catch(err => {
        console.error("Network error saving new master on server", err);
        showToast("Master enregistré localement mais non sauvegardé sur le serveur", "error");
      });

    // Reset form
    setNewMaster({
      testeur: "",
      idMaster: "",
      refProduitMaster: "",
      numSerieProduitMaster: "",
      refCarteMaster: "",
      numSerieCarteMaster: "",
      dateCreation: new Date().toLocaleDateString('fr-FR'),
      commentaire1: "",
      statut: "Active",
      commentaire2: "",
      verif: "OK"
    });
    
    // Go to list tab
    setActiveTab('testeurs');
  };

  // Delete master trigger (sets state for custom confirmation dialog)
  const handleDeleteMaster = (id: string) => {
    const item = masters.find(m => m.id === id);
    if (item) {
      setMasterToDelete(item);
    }
  };

  // Execute deletion after user confirms in our custom dialog
  const confirmDeleteMaster = () => {
    if (!masterToDelete) return;
    const idMaster = masterToDelete.idMaster;
    setMasters(prev => prev.filter(m => m.id !== masterToDelete.id));
    showToast(`Le Master ${idMaster} a été supprimé`, "success");
    addLog("Suppression Master", `Fiche Master supprimée - ID: ${masterToDelete.idMaster}, Testeur: ${masterToDelete.testeur}, Statut: ${masterToDelete.statut}`);
    setMasterToDelete(null);
  };

  // Edit master trigger (sets state for modification modal)
  const handleEditMaster = (item: MasterItem) => {
    setMasterToEdit({ ...item });
  };

  // Save modified master after user clicks save
  const saveEditMaster = () => {
    if (!masterToEdit) return;
    if (!masterToEdit.idMaster.trim()) {
      showToast("L'identification Master est obligatoire.", "error");
      return;
    }
    if (!masterToEdit.testeur.trim()) {
      showToast("Le testeur de destination est obligatoire.", "error");
      return;
    }
    
    setMasters(prev => prev.map(m => m.id === masterToEdit.id ? masterToEdit : m));
    showToast(`Master ${masterToEdit.idMaster} a été modifié avec succès`, "success");
    addLog("Modification Master", `Fiche Master modifiée - ID: ${masterToEdit.idMaster}, Testeur: ${masterToEdit.testeur}, Statut: ${masterToEdit.statut}, Vérif: ${masterToEdit.verif}`);
    setMasterToEdit(null);
  };

  // CSV Batch Importer logic
  const handleCsvImport = () => {
    if (!rawCsvPaste.trim()) {
      showToast("Veuillez coller des lignes de données semi-colonnes (;) d'abord.", "error");
      return;
    }

    const lines = rawCsvPaste.split("\n");
    let importedCount = 0;
    const newItems: MasterItem[] = [];

    lines.forEach(line => {
      if (!line.trim() || line.startsWith("TESTEURS") || line.startsWith("Suivi des")) return;
      const parts = line.split(";");
      if (parts.length >= 2) {
        const testeurVal = parts[0]?.trim();
        const idMasterVal = parts[1]?.trim() || "MASTER_TMP";
        
        // Skip header lines or descriptive labels
        if (!testeurVal || testeurVal === "TESTEURS" || idMasterVal === "Identification Master") {
          return;
        }

        // Check if already exists in active list
        if (masters.some(m => m.idMaster === idMasterVal) || newItems.some(n => n.idMaster === idMasterVal)) {
          return; // Skip duplicate
        }

        newItems.push({
          id: idMasterVal,
          testeur: testeurVal,
          idMaster: idMasterVal,
          refProduitMaster: parts[2]?.trim() || "",
          numSerieProduitMaster: parts[3]?.trim() || "",
          refCarteMaster: parts[4]?.trim() || "",
          numSerieCarteMaster: parts[5]?.trim() || "",
          dateCreation: parts[6]?.trim() || new Date().toLocaleDateString('fr-FR'),
          commentaire1: parts[7]?.trim() || "",
          statut: parts[8]?.trim() || "Active",
          commentaire2: parts[9]?.trim() || "",
          verif: parts[10]?.trim() || "OK"
        });
        importedCount++;
      }
    });

    if (newItems.length > 0) {
      setMasters(prev => [...newItems, ...prev]);
      showToast(`${importedCount} pièces masters importées avec succès !`, "success");
      addLog("Importation de masse (Excel/CSV)", `${importedCount} pièces masters importées dans la base de données.`);
      setRawCsvPaste("");
      setShowCsvImporter(false);
    } else {
      showToast("Aucun nouveau master importé (doublons ou format incorrect).", "error");
    }
  };


  // ----------------------------------------------------
  // Tab 3: CALCULATION STATE & FORMULAS
  // ----------------------------------------------------
  
  // Quick pre-filled capability calculation presets
  const [calcMasterId, setCalcMasterId] = useState(() => localStorage.getItem("mms_calc_master_id") || "");
  const [calcTester, setCalcTester] = useState(() => localStorage.getItem("mms_calc_tester") || "UFT SMOOVE RTS");
  const [calcCharacteristic, setCalcCharacteristic] = useState(() => localStorage.getItem("mms_calc_characteristic") || "Fréquence RF (MHz)");
  const [calcNominal, setCalcNominal] = useState<number>(() => {
    const v = localStorage.getItem("mms_calc_nominal");
    return v !== null ? parseFloat(v) : 433.42;
  });
  const [calcLsl, setCalcLsl] = useState<number>(() => {
    const v = localStorage.getItem("mms_calc_lsl");
    return v !== null ? parseFloat(v) : 433.22;
  });
  const [calcUsl, setCalcUsl] = useState<number>(() => {
    const v = localStorage.getItem("mms_calc_usl");
    return v !== null ? parseFloat(v) : 433.62;
  });
  const [calcOperator, setCalcOperator] = useState(() => localStorage.getItem("mms_calc_operator") || "BEN MANSOUR Samir");
  
  // State for deletion confirmation modal
  const [masterToDelete, setMasterToDelete] = useState<MasterItem | null>(null);
  
  // State for modification modal
  const [masterToEdit, setMasterToEdit] = useState<MasterItem | null>(null);
  
  // User Session Management (V9: restore from opaque token + identity cache)
  const [session, setSession] = useState<UserSession | null>(() => {
    const token = (typeof localStorage !== "undefined") ? localStorage.getItem("mms_token") : null;
    const savedIdentity = (typeof localStorage !== "undefined") ? localStorage.getItem("mms_user_session") : null;

    let role: string | undefined;
    let matricule: string | undefined;
    let name: string | undefined;

    if (savedIdentity) {
      try {
        const parsed = JSON.parse(savedIdentity);
        role = parsed.role;
        matricule = parsed.matricule;
        name = parsed.name;
      } catch (e) {
        console.error("Error reading saved user session", e);
      }
    }

    if (token) {
      const claims = decodeV9Token(token);
      if (claims && claims.role && claims.matricule) {
        return {
          role: claims.role,
          matricule: claims.matricule,
          name: name || "",
          token
        };
      }
      localStorage.removeItem("mms_token");
    }

    if (role && matricule) {
      return { role, matricule, name: name || "", token: "" };
    }

    return null;
  });

  // Authorized Technicians State: initialize from defaults, then fetch from server
  const [technicians, setTechnicians] = useState<{ matricule: string; name: string }[]>(() => {
    const DEFAULT_TECHS = [
      { matricule: "MTR01", name: "BEN MANSOUR Samir" },
      { matricule: "MTR02", name: "DURAND Nicolas" }
    ];
    return DEFAULT_TECHS;
  });

  // Fetch technicians from server on mount
  useEffect(() => {
    fetch("/api/technicians", { headers: authHeaders() })
      .then(res => {
        if (!res.ok) throw new Error("Server response was not ok");
        return res.json();
      })
      .then(data => {
        if (data && Array.isArray(data) && data.length > 0) {
          setTechnicians(data);
        }
      })
      .catch(err => {
        console.error("Failed to load technicians from server, using local fallback", err);
        const saved = localStorage.getItem("mms_technicians");
        if (saved) {
          try {
            const parsed: { matricule: string; name: string }[] = JSON.parse(saved);
            const DEFAULT_TECHS = [
              { matricule: "MTR01", name: "BEN MANSOUR Samir" },
              { matricule: "MTR02", name: "DURAND Nicolas" }
            ];
            const merged = [...DEFAULT_TECHS];
            for (const t of parsed) {
              if (!merged.some(d => d.matricule === t.matricule)) {
                merged.push(t);
              }
            }
            setTechnicians(merged);
          } catch (e) {
            console.error("Error reading saved technicians", e);
          }
        }
      });
  }, []);

  // Save technicians to localStorage as cache
  useEffect(() => {
    localStorage.setItem("mms_technicians", JSON.stringify(technicians));
  }, [technicians]);

  // Logs / Change History State
  const [logs, setLogs] = useState<HistoryLog[]>(() => {
    const saved = localStorage.getItem("mms_change_logs");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error("Error reading saved logs", e);
      }
    }
    return [];
  });

  // Helper function to record a log
  const addLog = (action: string, details: string) => {
    const newLog: HistoryLog = {
      id: `log_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
      timestamp: new Date().toLocaleString('fr-FR'),
      user: session ? `${session.name} (${session.matricule})` : "Système",
      role: session?.role || 'technician',
      action,
      details
    };
    setLogs(prev => {
      const updated = [newLog, ...prev];
      localStorage.setItem("mms_change_logs", JSON.stringify(updated));
      return updated;
    });
  };

  // Admin form fields for adding a technician
  const [newTechMatricule, setNewTechMatricule] = useState("");
  const [newTechName, setNewTechName] = useState("");
  const [logSearchQuery, setLogSearchQuery] = useState("");

  // Confirmation & Panel states
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [techToDelete, setTechToDelete] = useState<string | null>(null);

  // Filtered change logs
  const filteredLogs = useMemo(() => {
    if (!logSearchQuery.trim()) return logs;
    const q = logSearchQuery.toLowerCase();
    return logs.filter(log => 
      log.action.toLowerCase().includes(q) ||
      log.details.toLowerCase().includes(q) ||
      log.user.toLowerCase().includes(q)
    );
  }, [logs, logSearchQuery]);

  const handleAddTechnician = (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTechMatricule.trim()) {
      showToast("Matricule obligatoire.", "error");
      return;
    }
    if (!newTechName.trim()) {
      showToast("Nom complet obligatoire.", "error");
      return;
    }
    const mat = newTechMatricule.trim().toUpperCase();
    if (technicians.some(t => t.matricule === mat)) {
      showToast(`Le technicien avec le matricule ${mat} existe déjà.`, "error");
      return;
    }
    const created = { matricule: mat, name: newTechName.trim() };
    setTechnicians(prev => [...prev, created]);
    addLog("Ajout Technicien", `Nouveau compte de technicien enregistré : ${created.name} (${created.matricule})`);
    showToast(`Technicien ${created.name} enregistré avec succès !`, "success");
    setNewTechMatricule("");
    setNewTechName("");

    const sessionToken = session?.token || localStorage.getItem("mms_token") || API_TOKEN;
    fetch("/api/technicians", {
      method: "POST",
      headers: sessionToken
        ? { "Content-Type": "application/json", Authorization: `Bearer ${sessionToken}` }
        : authHeaders(),
      body: JSON.stringify(created)
    }).then(res => res.ok ? res.json() : Promise.reject(new Error("Server write failed")))
      .then(data => {
        if (!data.success) {
          console.error("Failed to save technician on server:", data.error);
          showToast("Technicien enregistré localement mais non sauvegardé sur le serveur", "error");
        }
      })
      .catch(err => {
        console.error("Network error saving technician on server", err);
        showToast("Technicien enregistré localement mais non sauvegardé sur le serveur", "error");
      });
  };

  const handleDeleteTechnician = (matricule: string) => {
    const tech = technicians.find(t => t.matricule === matricule);
    if (!tech) return;
    setTechnicians(prev => prev.filter(t => t.matricule !== matricule));
    addLog("Suppression Technicien", `Compte du technicien supprimé : ${tech.name} (${tech.matricule})`);
    showToast(`Technicien ${tech.name} supprimé.`, "info");
    setTechToDelete(null);

    const sessionToken = session?.token || localStorage.getItem("mms_token") || API_TOKEN;
    fetch(`/api/technicians/${encodeURIComponent(matricule)}`, {
      method: "DELETE",
      headers: sessionToken
        ? { "Content-Type": "application/json", Authorization: `Bearer ${sessionToken}` }
        : authHeaders()
    }).catch(err => {
      console.error("Network error deleting technician on server", err);
      showToast("Technicien supprimé localement mais non supprimé sur le serveur", "error");
    });
  };

  const handleClearLogs = () => {
    setLogs([]);
    localStorage.removeItem("mms_change_logs");
    showToast("Historique vidé !", "success");
    setShowClearConfirm(false);
  };

  const handleExportLogs = () => {
    if (logs.length === 0) {
      showToast("Aucun historique à exporter.", "error");
      return;
    }
    const header = "HISTORIQUE D'AUDIT ET CHANGEMENTS - ACTIA TUNISIE\n";
    const subheader = `Généré le : ${new Date().toLocaleString('fr-FR')}\n`;
    const separator = "=".repeat(80) + "\n";
    
    const logLines = logs.map(log => {
      return `[${log.timestamp}] ACTION: ${log.action}\nPar : ${log.user}\nDétails: ${log.details}\n${"-".repeat(60)}`;
    }).join("\n\n");
    
    const content = header + subheader + separator + "\n" + logLines;
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `historique_actia_${new Date().toISOString().slice(0, 10)}.txt`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast("Historique exporté avec succès en fichier TXT !", "success");
  };

  // Login form fields state
  const [loginRole, setLoginRole] = useState<'admin' | 'technician'>('technician');
  const [loginMatricule, setLoginMatricule] = useState("");
  const [loginName, setLoginName] = useState("");
  const [loginAdminCode, setLoginAdminCode] = useState("");

  // Keep track of session in localStorage and update operator
  useEffect(() => {
    if (session) {
      if (session.token) {
        localStorage.setItem("mms_token", session.token);
      }
      localStorage.setItem("mms_user_session", JSON.stringify({
        role: session.role,
        matricule: session.matricule,
        name: session.name
      }));
      setCalcOperator(`${session.name} (${session.matricule})`);
    } else {
      localStorage.removeItem("mms_token");
      localStorage.removeItem("mms_user_session");
    }
  }, [session]);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    
    let token: string | null = null;
    const userMatricule = loginMatricule.trim().toUpperCase();
    
    if (loginRole === 'technician') {
      if (!userMatricule) {
        showToast("Le matricule est obligatoire pour la session technicien.", "error");
        return;
      }
      let matchedTech = technicians.find(
        t => t.matricule.trim().toUpperCase() === userMatricule
      );
      if (!matchedTech) {
        const newName = loginName.trim();
        if (!newName) {
          showToast("Le nom du technicien est obligatoire pour une nouvelle inscription.", "error");
          return;
        }
        const created = { matricule: userMatricule, name: newName };
        setTechnicians(prev => [...prev, created]);
        matchedTech = created;
        showToast(`Nouveau technicien enregistré : ${created.name} (${created.matricule})`, "success");

        const sessionToken = session?.token || localStorage.getItem("mms_token") || API_TOKEN;
        fetch("/api/technicians", {
          method: "POST",
          headers: sessionToken
            ? { "Content-Type": "application/json", Authorization: `Bearer ${sessionToken}` }
            : authHeaders(),
          body: JSON.stringify(created)
        }).catch(err => {
          console.error("Network error saving new technician on server", err);
        });
      }
      
      try {
        const res = await fetch("/api/auth/login", {
          method: "POST",
          headers: authHeaders(),
          body: JSON.stringify({
            matricule: userMatricule,
            role: "technician"
          })
        });
        if (!res.ok) throw new Error("Technician login failed");
        const data = await res.json();
        token = data.token;
      } catch {
        showToast("Échec de l'authentification auprès du serveur.", "error");
        return;
      }
      
      const newSession: UserSession = {
        role: 'technician',
        matricule: matchedTech.matricule,
        name: matchedTech.name,
        token: token!
      };
      setSession(newSession);
      showToast(`Session Technicien activée : ${newSession.name} (${newSession.matricule})`, "success");
    } else {
      if (!userMatricule) {
        showToast("Le matricule Admin est obligatoire.", "error");
        return;
      }
      if (!loginAdminCode.trim()) {
        showToast("Le code d'accès Admin est obligatoire.", "error");
        return;
      }
      
      try {
        const res = await fetch("/api/auth/login", {
          method: "POST",
          headers: authHeaders(),
          body: JSON.stringify({
            matricule: userMatricule,
            code: loginAdminCode,
            role: "admin"
          })
        });
        if (!res.ok) throw new Error("Admin login failed");
        const data = await res.json();
        token = data.token;
      } catch {
        showToast("Code d'accès Admin incorrect ou erreur serveur.", "error");
        return;
      }
      
      const newSession: UserSession = {
        role: 'admin',
        matricule: userMatricule,
        name: loginName.trim() || "ADMIN METROLOGIE",
        token: token!
      };
      setSession(newSession);
      showToast(`Session Administrateur activée ! Accès complet autorisé.`, "success");
    }
  };

  const handleLogout = () => {
    setSession(null);
    localStorage.removeItem("mms_token");
    localStorage.removeItem("mms_user_session");
    setLoginMatricule("");
    setLoginName("");
    setLoginAdminCode("");
    showToast("Déconnexion réussie. Session fermée.", "info");
  };
  
  // Local background executable bridge integration
  const [useExternalCalc, setUseExternalCalc] = useState(false);
  const [calcEngine, setCalcEngine] = useState<'web' | 'local' | 'ai'>(() => (localStorage.getItem("mms_calc_engine") as 'web' | 'local' | 'ai') || 'web');
  const [calcIsSc, setCalcIsSc] = useState<'NON' | 'OUI'>(() => (localStorage.getItem("mms_calc_is_sc") as 'NON' | 'OUI') || 'NON');
  const [isAiCalculating, setIsAiCalculating] = useState(false);
  const [aiStatsResult, setAiStatsResult] = useState<{
    mean: number;
    stdDev: number;
    cp: number;
    cpk: number;
    cpkType: string;
    range: number;
    repeatability: number;
    notes: string;
    n: number;
    outliers?: number[];
    diagnostic?: string;
  } | null>(null);
  const [externalCalcMethod, setExternalCalcMethod] = useState<'api' | 'protocol'>('api');
  const [localBridgeUrl, setLocalBridgeUrl] = useState("http://localhost:5001/calculate");
  const [localExePath, setLocalExePath] = useState("C:\\ACTIA\\Metrologie\\calcpk.exe");
  const [isLocalTesting, setIsLocalTesting] = useState(false);
  const [isLocalCalculating, setIsLocalCalculating] = useState(false);
  const [localBridgeStatus, setLocalBridgeStatus] = useState<'idle' | 'success' | 'error'>('idle');
  const [bridgeScriptLanguage, setBridgeScriptLanguage] = useState<'python' | 'node'>('python');
  const [externalStatsResult, setExternalStatsResult] = useState<{
    mean: number;
    stdDev: number;
    cp: number;
    cpk: number;
    cpkType: string;
    range: number;
    repeatability: number;
    notes: string;
    n: number;
  } | null>(null);
  
  // Master Log File Import state and logic
  const [isDraggingLog, setIsDraggingLog] = useState(false);
  const [logImportDetails, setLogImportDetails] = useState<{
    fileName: string;
    lineCount: number;
    valuesFound: number[];
    nominal?: number;
    lsl?: number;
    usl?: number;
    characteristic?: string;
    files?: Array<{
      fileName: string;
      lineCount: number;
      valuesFound: number[];
      nominal?: number;
      lsl?: number;
      usl?: number;
      characteristic?: string;
    }>;
  } | null>(null);

  const readAndParseLogFile = async (file: File): Promise<{
    fileName: string;
    lineCount: number;
    valuesFound: number[];
    nominal?: number;
    lsl?: number;
    usl?: number;
    characteristic?: string;
  }> => {
    try {
      let text = "";
      try {
        const arrayBuffer = await file.arrayBuffer();
        const utf8Decoder = new TextDecoder('utf-8', { fatal: true });
        text = utf8Decoder.decode(arrayBuffer);
      } catch (e) {
        const arrayBuffer = await file.arrayBuffer();
        const winDecoder = new TextDecoder('windows-1252');
        text = winDecoder.decode(arrayBuffer);
      }

      const lines = text.split("\n");
      let nominalVal: number | undefined;
      let lslVal: number | undefined;
      let uslVal: number | undefined;
      let characteristicVal: string | undefined;

      lines.forEach(line => {
        const lower = line.toLowerCase();
        if (lower.includes("caracteristique") || lower.includes("characteristic") || lower.includes("test name") || lower.includes("test_name")) {
          const parts = line.split(/[;=:]/);
          if (parts.length > 1) {
            characteristicVal = parts[1].trim();
          }
        }

        const matchLimit = (label: string) => {
          if (lower.includes(label)) {
            const parts = line.split(/[;=:\s]+/);
            for (let i = 0; i < parts.length; i++) {
              if (parts[i].toLowerCase().includes(label)) {
                const valPart = parts[i+1] || parts[i].split(/[;=:]/)[1];
                if (valPart) {
                  const parsed = parseFloat(valPart.replace(',', '.').replace(/[^\d.-]/g, ''));
                  if (!isNaN(parsed)) return parsed;
                }
              }
            }
          }
          return undefined;
        };

        const nVal = matchLimit("nominal") || matchLimit("target") || matchLimit("centre");
        if (nVal !== undefined) nominalVal = nVal;

        const minVal = matchLimit("lsl") || matchLimit("min") || matchLimit("limit_inf") || matchLimit("inf");
        if (minVal !== undefined) lslVal = minVal;

        const maxVal = matchLimit("usl") || matchLimit("max") || matchLimit("limit_sup") || matchLimit("sup");
        if (maxVal !== undefined) uslVal = maxVal;
      });

      // Match float values (like 433.42, 5.01, etc.)
      const regex = /-?\b\d+[\.,]\d+\b/g;
      const matches = text.match(regex) || [];
      let parsedValues = matches
        .map(m => parseFloat(m.replace(',', '.')))
        .filter(v => !isNaN(v));

      if (parsedValues.length === 0) {
        lines.forEach(line => {
          const trimmed = line.trim().replace(',', '.');
          const num = parseFloat(trimmed);
          if (trimmed && !isNaN(num) && trimmed.match(/^-?\d+$/)) {
            parsedValues.push(num);
          }
        });
      }

      parsedValues = parsedValues.filter(val => {
        if (val >= 2000 && val <= 2100) return false; // filter out year numbers
        return true;
      });

      return {
        fileName: file.name,
        lineCount: lines.length,
        valuesFound: parsedValues,
        nominal: nominalVal,
        lsl: lslVal,
        usl: uslVal,
        characteristic: characteristicVal
      };
    } catch (err) {
      console.error(`Erreur lecture fichier ${file.name}:`, err);
      return {
        fileName: file.name,
        lineCount: 0,
        valuesFound: []
      };
    }
  };

  const handleMultipleLogFiles = async (files: File[]) => {
    if (files.length === 0) return;

    if (files.length < 10) {
      showToast("Veuillez sélectionner au minimum 10 fichiers pour lancer le calcul.", "error");
      return;
    }

    try {
      const parsedResults = await Promise.all(files.map(file => readAndParseLogFile(file)));
      const validResults = parsedResults.filter(r => r.valuesFound.length > 0);

      if (validResults.length === 0) {
        showToast("Aucune valeur numérique détectée dans les fichiers sélectionnés.", "error");
        return;
      }

      // Aggregate all measurements
      const allValues: number[] = [];
      validResults.forEach(r => {
        allValues.push(...r.valuesFound);
      });

      // Find spec metadata from first available file that has them
      const firstNominal = validResults.find(r => r.nominal !== undefined)?.nominal;
      const firstLsl = validResults.find(r => r.lsl !== undefined)?.lsl;
      const firstUsl = validResults.find(r => r.usl !== undefined)?.usl;
      const firstCharacteristic = validResults.find(r => r.characteristic !== undefined)?.characteristic;

      setLogImportDetails({
        fileName: validResults.length === 1 ? validResults[0].fileName : `${validResults.length} fichiers log sélectionnés`,
        lineCount: validResults.reduce((acc, curr) => acc + curr.lineCount, 0),
        valuesFound: allValues,
        nominal: firstNominal,
        lsl: firstLsl,
        usl: firstUsl,
        characteristic: firstCharacteristic,
        files: validResults
      });

      showToast(`${validResults.length} fichier(s) traité(s) avec succès ! ${allValues.length} valeurs importées.`, "success");
    } catch (err) {
      console.error(err);
      showToast("Erreur lors du traitement des fichiers logs.", "error");
    }
  };

  const handleLogFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!e.target.files) return;
    const files = Array.from(e.target.files) as File[];
    handleMultipleLogFiles(files);
  };

  const handleLogFileDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDraggingLog(false);
    if (!e.dataTransfer.files) return;
    const files = Array.from(e.dataTransfer.files) as File[];
    handleMultipleLogFiles(files);
  };

  const handleInjectLogValues = (mode: 'replace' | 'append') => {
    if (!logImportDetails) return;
    const formattedValues = logImportDetails.valuesFound.join("\n");
    if (mode === 'replace') {
      setMeasurementsInput(formattedValues);
      showToast("Relevés existants remplacés par le log.", "success");
    } else {
      setMeasurementsInput(prev => prev ? `${prev}\n${formattedValues}` : formattedValues);
      showToast("Valeurs du log ajoutées aux relevés existants.", "success");
    }

    if (logImportDetails.nominal !== undefined) setCalcNominal(logImportDetails.nominal);
    if (logImportDetails.lsl !== undefined) setCalcLsl(logImportDetails.lsl);
    if (logImportDetails.usl !== undefined) setCalcUsl(logImportDetails.usl);
    if (logImportDetails.characteristic) setCalcCharacteristic(logImportDetails.characteristic);

    setLogImportDetails(null);
  };

  // Measurements list
  const [measurementsInput, setMeasurementsInput] = useState<string>(() => {
    return localStorage.getItem("mms_measurements_input") || "433.41\n433.45\n433.39\n433.43\n433.42\n433.40\n433.44\n433.41\n433.43\n433.42\n433.41\n433.45\n433.38\n433.42\n433.43";
  });

  // Save calculation inputs to localStorage whenever they change
  useEffect(() => {
    localStorage.setItem("mms_calc_master_id", calcMasterId);
    localStorage.setItem("mms_calc_tester", calcTester);
    localStorage.setItem("mms_calc_characteristic", calcCharacteristic);
    localStorage.setItem("mms_calc_nominal", String(calcNominal));
    localStorage.setItem("mms_calc_lsl", String(calcLsl));
    localStorage.setItem("mms_calc_usl", String(calcUsl));
    localStorage.setItem("mms_calc_operator", calcOperator);
    localStorage.setItem("mms_calc_engine", calcEngine);
    localStorage.setItem("mms_calc_is_sc", calcIsSc);
    localStorage.setItem("mms_measurements_input", measurementsInput);
  }, [calcMasterId, calcTester, calcCharacteristic, calcNominal, calcLsl, calcUsl, calcOperator, calcEngine, calcIsSc, measurementsInput]);

  // Individual interactive list of measurement numbers
  const measurementsArray = useMemo(() => {
    return measurementsInput
      .split(/[\n,; \t]+/)
      .map(v => parseFloat(v.replace(',', '.')))
      .filter(v => !isNaN(v));
  }, [measurementsInput]);

  // Change the tester for calculation: automatically clear any previously loaded
  // log/relevé file and measurements so a stale file from another tester is never reused.
  const handleCalcTesterChange = (value: string) => {
    setCalcTester(value);
    setLogImportDetails(null);
    setMeasurementsInput("");
    showToast("Testeur modifié : anciens fichiers/relevés chargés vidés.", "info");
  };

  // Update inputs when a master is selected for calculation
  const handleSelectMasterForCalculation = (item: MasterItem) => {
    setCalcMasterId(item.idMaster);
    setCalcTester(item.testeur);
    setLogImportDetails(null);
    
    // Choose appropriate default nominals based on product references
    if (item.testeur.toLowerCase().includes("telis")) {
      setCalcCharacteristic("Tension d'alimentation (V)");
      setCalcNominal(5.0);
      setCalcLsl(4.75);
      setCalcUsl(5.25);
      setMeasurementsInput("4.98\n5.02\n5.05\n4.97\n5.00\n5.01\n4.99\n5.03\n5.02\n4.96\n5.01\n5.00\n4.99\n5.02\n5.01");
    } else if (item.testeur.toLowerCase().includes("smoove") || item.testeur.toLowerCase().includes("rts")) {
      setCalcCharacteristic("Fréquence RF (MHz)");
      setCalcNominal(433.42);
      setCalcLsl(433.22);
      setCalcUsl(433.62);
      setMeasurementsInput("433.41\n433.45\n433.39\n433.43\n433.42\n433.40\n433.44\n433.41\n433.43\n433.42\n433.41\n433.45\n433.38\n433.42\n433.43");
    } else if (item.testeur.toLowerCase().includes("conso")) {
      setCalcCharacteristic("Consommation Courant (mA)");
      setCalcNominal(120.0);
      setCalcLsl(110.0);
      setCalcUsl(130.0);
      setMeasurementsInput("121.2\n118.5\n119.8\n122.1\n120.4\n117.9\n120.1\n121.5\n119.2\n120.8\n121.1\n118.9\n119.5\n122.3\n120.2");
    } else if (item.testeur.toLowerCase().includes("anemometre") || item.testeur.toLowerCase().includes("eolis")) {
      setCalcCharacteristic("Seuil Anémomètre (m/s)");
      setCalcNominal(12.0);
      setCalcLsl(10.5);
      setCalcUsl(13.5);
      setMeasurementsInput("11.8\n12.1\n11.9\n12.4\n12.0\n11.7\n12.2\n12.1\n11.8\n12.0\n11.9\n12.3\n12.0\n11.6\n12.1");
    } else {
      // Default general preset
      setCalcCharacteristic("Dimension critique (mm)");
      setCalcNominal(10.0);
      setCalcLsl(9.8);
      setCalcUsl(10.2);
      setMeasurementsInput("10.01\n10.02\n9.99\n10.03\n10.00\n10.01\n9.98\n10.02\n10.01\n10.00\n10.02\n9.99\n10.01\n10.00\n10.03");
    }

    showToast(`Master ${item.idMaster} sélectionné. Paramètres pré-remplis !`, "info");
    setActiveTab('calcul');
  };

  // Presets load helper
  const loadPreset = (type: '5v' | 'conso' | 'rf' | 'dim') => {
    switch(type) {
      case '5v':
        setCalcCharacteristic("Tension Régulée (V)");
        setCalcNominal(5.00);
        setCalcLsl(4.85);
        setCalcUsl(5.15);
        setMeasurementsInput("4.98\n5.01\n5.02\n4.97\n4.99\n5.00\n5.03\n4.96\n5.01\n5.00\n5.02\n4.98\n4.99\n5.01\n5.00");
        showToast("Preset 5V chargé", "info");
        break;
      case 'conso':
        setCalcCharacteristic("Courant de repos (mA)");
        setCalcNominal(24.0);
        setCalcLsl(21.0);
        setCalcUsl(27.0);
        setMeasurementsInput("24.2\n23.8\n24.5\n23.1\n24.0\n24.9\n23.7\n24.3\n24.1\n23.9\n24.4\n23.6\n24.0\n24.2\n23.8");
        showToast("Preset Courant loaded", "info");
        break;
      case 'rf':
        setCalcCharacteristic("Fréquence Smoove RTS (MHz)");
        setCalcNominal(433.42);
        setCalcLsl(433.22);
        setCalcUsl(433.62);
        setMeasurementsInput("433.41\n433.43\n433.40\n433.42\n433.45\n433.41\n433.43\n433.42\n433.39\n433.44\n433.41\n433.42\n433.43\n433.40\n433.42");
        showToast("Preset RF loaded", "info");
        break;
      case 'dim':
        setCalcCharacteristic("Entrefer Électroaimant (mm)");
        setCalcNominal(1.50);
        setCalcLsl(1.35);
        setCalcUsl(1.65);
        setMeasurementsInput("1.48\n1.51\n1.47\n1.53\n1.50\n1.49\n1.52\n1.48\n1.51\n1.50\n1.49\n1.52\n1.46\n1.51\n1.50");
        showToast("Preset Dimension loaded", "info");
        break;
    }
  };

  // Core statistical math calculations
  const statsResult = useMemo(() => {
    const arr = measurementsArray;
    const n = arr.length;
    if (n === 0) return null;

    const minVal = Math.min(...arr);
    const maxVal = Math.max(...arr);
    const rangeVal = maxVal - minVal;
    
    // Average
    const sum = arr.reduce((acc, curr) => acc + curr, 0);
    const meanVal = sum / n;

    // Standard deviation (sample standard deviation formula N-1)
    let variance = 0;
    if (n > 1) {
      const sumSqDiff = arr.reduce((acc, curr) => acc + Math.pow(curr - meanVal, 2), 0);
      variance = sumSqDiff / (n - 1);
    }
    const stdDevVal = Math.sqrt(variance);

    // Cp and Cpk
    let cpVal = 0;
    let cpkVal = 0;
    let cpkMinType: 'LSL' | 'USL' | 'Perfect' = 'Perfect';

    if (stdDevVal > 0) {
      cpVal = (calcUsl - calcLsl) / (6 * stdDevVal);
      const cpkUpper = (calcUsl - meanVal) / (3 * stdDevVal);
      const cpkLower = (meanVal - calcLsl) / (3 * stdDevVal);
      if (cpkLower < cpkUpper) {
        cpkVal = cpkLower;
        cpkMinType = 'LSL';
      } else {
        cpkVal = cpkUpper;
        cpkMinType = 'USL';
      }
    } else {
      cpVal = Infinity;
      cpkVal = Infinity;
    }

    // Repeatability (Industrial 6-sigma variation)
    const repeatabilityVal = 6 * stdDevVal;

    // Evaluate capability status
    let statusLabel = "Non capable (Inacceptable)";
    let statusColor = "red";
    const thresholdCpk = calcIsSc === 'OUI' ? 1.67 : 1.33;
    const warningCpk = calcIsSc === 'OUI' ? 1.33 : 1.0;
    if (cpkVal >= thresholdCpk) {
      statusLabel = "Hautement Capable (Classe Mondiale)";
      statusColor = "green";
    } else if (cpkVal >= warningCpk) {
      statusLabel = "Capable (Acceptable sous contrôle)";
      statusColor = "yellow";
    }

    return {
      n,
      min: minVal,
      max: maxVal,
      range: rangeVal,
      mean: meanVal,
      stdDev: stdDevVal,
      cp: cpVal,
      cpk: cpkVal,
      cpkType: cpkMinType,
      repeatability: repeatabilityVal,
      statusLabel,
      statusColor
    };
  }, [measurementsArray, calcLsl, calcUsl]);

  // Test connection to the local background agent/bridge
  const handleTestLocalBridge = async () => {
    setIsLocalTesting(true);
    setLocalBridgeStatus('idle');
    try {
      const response = await fetch(localBridgeUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          test: true,
          exePath: localExePath,
          testerName: calcTester,
          characteristic: calcCharacteristic,
          operator: calcOperator,
          nominal: calcNominal,
          lsl: calcLsl,
          usl: calcUsl,
          measurements: measurementsArray.slice(0, 3) // short array for test
        })
      });
      if (response.ok) {
        setLocalBridgeStatus('success');
        showToast("Connexion réussie avec le démon local ! Pont actif.", "success");
      } else {
        setLocalBridgeStatus('error');
        showToast("Le démon local a répondu avec une erreur.", "error");
      }
    } catch (e) {
      console.error(e);
      setLocalBridgeStatus('error');
      showToast("Impossible de contacter le démon local. Vérifiez qu'il est lancé sur le port 5001.", "error");
    } finally {
      setIsLocalTesting(false);
    }
  };

  // Run the calculations via the local executable background daemon
  const handleExecuteExternalCalc = async () => {
    if (measurementsArray.length === 0) {
      showToast("Veuillez d'abord saisir ou importer des mesures.", "error");
      return;
    }
    setIsLocalCalculating(true);
    try {
      showToast("Appel de l'exécutable local en arrière-plan...", "info");
      const response = await fetch(localBridgeUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          exePath: localExePath,
          testerName: calcTester,
          characteristic: calcCharacteristic,
          operator: calcOperator,
          nominal: calcNominal,
          lsl: calcLsl,
          usl: calcUsl,
          measurements: measurementsArray
        })
      });
      
      if (response.ok) {
        const data = await response.json();
        if (data.success) {
          // Extract metrics returned by user's exe
          setExternalStatsResult({
            mean: data.mean ?? statsResult?.mean ?? 0,
            stdDev: data.stdDev ?? statsResult?.stdDev ?? 0,
            cp: data.cp ?? statsResult?.cp ?? 0,
            cpk: data.cpk ?? statsResult?.cpk ?? 0,
            cpkType: data.cpkType ?? statsResult?.cpkType ?? 'LSL',
            range: data.range ?? statsResult?.range ?? 0,
            repeatability: data.repeatability ?? statsResult?.repeatability ?? 0,
            notes: data.notes ?? `Calculé par l'exécutable local: ${localExePath.split('\\').pop()}`,
            n: measurementsArray.length
          });
          showToast("Calculs métrologiques exécutés avec succès par votre application locale !", "success");
        } else {
          throw new Error(data.error || "Erreur retournée par le pont local.");
        }
      } else {
        throw new Error("Le serveur local a renvoyé un code HTTP " + response.status);
      }
    } catch (e: any) {
      console.error(e);
      
      // Fallback with notification: Let them simulate successful execution if they want to try the UI flow
      showToast(`Échec du pont local : ${e.message}. Simulation des résultats locaux activée.`, "info");
      
      // Generate simulated response so the dashboard works even if they haven't launched the daemon yet!
      if (statsResult) {
        setExternalStatsResult({
          mean: statsResult.mean,
          stdDev: statsResult.stdDev,
          cp: statsResult.cp,
          cpk: statsResult.cpk,
          cpkType: statsResult.cpkType,
          range: statsResult.range,
          repeatability: statsResult.repeatability,
          notes: `[SIMULÉ - PONT ABSENT] Serait calculé en appelant l'exécutable : ${localExePath.split('\\').pop()}`,
          n: statsResult.n
        });
      }
    } finally {
      setIsLocalCalculating(false);
    }
  };

  // Run the calculations via the server-side Gemini API
  const handleExecuteAiCalc = async () => {
    if (measurementsArray.length === 0) {
      showToast("Veuillez d'abord saisir ou importer des mesures.", "error");
      return;
    }
    setIsAiCalculating(true);
    try {
      showToast("Appel de l'intelligence artificielle Gemini en cours...", "info");
      const response = await fetch("/api/gemini/calculate-capability", {
        method: "POST",
        headers: authHeaders(),
        body: JSON.stringify({
          mesuresBrutes: measurementsArray,
          lsl: calcLsl,
          usl: calcUsl,
          nominal: calcNominal,
          estSC: calcIsSc === 'OUI'
        })
      });

      if (!response.ok) {
        throw new Error("Erreur de communication avec le serveur backend.");
      }

      const data = await response.json();
      if (data.success && data.result) {
        const res = data.result;
        
        setAiStatsResult({
          mean: res.details_metriques.moyenne_av,
          stdDev: res.details_metriques.ecart_type_std,
          cp: res.details_metriques.potentiel_cp,
          cpk: res.details_metriques.ajuste_cpk,
          cpkType: res.outliers_detectes.length > 0 ? "Ajusté (Outliers filtrés)" : "LSL",
          range: res.details_metriques.etendue_r,
          repeatability: res.details_metriques.repetabilite_ev,
          notes: `[Intelligence Artificielle] ${res.outliers_detectes.length} valeur(s) aberrante(s) identifiée(s) et nettoyée(s).`,
          n: res.mesures_nettoyees.length,
          outliers: res.outliers_detectes,
          diagnostic: res.explication_diagnostic
        });

        if (res.outliers_detectes.length > 0) {
          showToast(`Analyse terminée ! ${res.outliers_detectes.length} valeur(s) aberrante(s) détectée(s) : [${res.outliers_detectes.join(", ")}].`, "info");
        } else {
          showToast("Analyse terminée ! Aucune valeur aberrante détectée.", "success");
        }
      } else {
        throw new Error(data.error || "Réponse invalide du serveur de calcul IA.");
      }
    } catch (err: any) {
      console.error(err);
      showToast(`Échec du calcul IA : ${err.message}`, "error");
    } finally {
      setIsAiCalculating(false);
    }
  };

  // Launch local app using custom protocol URL scheme (Windows Protocol Handler)
  const handleTriggerCustomProtocol = () => {
    const dataStr = encodeURIComponent(JSON.stringify({
      testerName: calcTester,
      characteristic: calcCharacteristic,
      operator: calcOperator,
      nominal: calcNominal,
      lsl: calcLsl,
      usl: calcUsl,
      measurements: measurementsArray
    }));
    const uri = `actiacalc://run?exe=${encodeURIComponent(localExePath)}&data=${dataStr}`;
    showToast("Appel du protocole personnalisé actiacalc://...", "info");
    window.location.href = uri;
  };

  // Unified active statistics used in report and charts
  const activeStats = useMemo(() => {
    if (calcEngine === 'ai') {
      return aiStatsResult;
    }
    const base = calcEngine === 'local' ? externalStatsResult : statsResult;
    if (!base) return null;

    let statusLabel = "Non capable (Inacceptable)";
    let statusColor = "red";
    const thresholdCpk = calcIsSc === 'OUI' ? 1.67 : 1.33;
    const warningCpk = calcIsSc === 'OUI' ? 1.33 : 1.0;

    if (base.cpk >= thresholdCpk) {
      statusLabel = "Hautement Capable (Classe Mondiale)";
      statusColor = "green";
    } else if (base.cpk >= warningCpk) {
      statusLabel = "Capable (Acceptable sous contrôle)";
      statusColor = "yellow";
    }

    return {
      ...base,
      statusLabel,
      statusColor
    };
  }, [calcEngine, externalStatsResult, statsResult, aiStatsResult, calcIsSc]);

  // Generate synthetic test measurements centered around nominal with random noise
  const generateSyntheticData = () => {
    const count = 15;
    const values: number[] = [];
    const toleranceRange = calcUsl - calcLsl;
    // We want a standard deviation of about toleranceRange / 8 to ensure capability is reasonable (Cpk around 1.3)
    const targetSigma = toleranceRange / 9;

    for (let i = 0; i < count; i++) {
      // Box-Muller transform for normal distribution random variable
      const u1 = Math.random() || 0.0001; 
      const u2 = Math.random();
      const z0 = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);
      
      // Scale with sigma and add to nominal or slightly off-nominal
      const val = calcNominal + (z0 * targetSigma);
      values.push(parseFloat(val.toFixed(3)));
    }

    setMeasurementsInput(values.join("\n"));
    showToast("15 essais de mesures générés aléatoirement selon la loi normale !", "success");
  };

  // Print friendly view trigger
  const handlePrint = () => {
    addLog("Impression du rapport", `Impression d'un rapport de capabilité pour le Master ${calcMasterId || "Manuel"} - Caractéristique: ${calcCharacteristic}`);
    window.print();
  };

  // Professional PDF export handler with high-res DOM canvas capture
  const handleExportPDF = async () => {
    if (!activeStats) {
      showToast("Aucune donnée disponible pour exporter le rapport.", "error");
      return;
    }

    setIsExportingPDF(true);
    addLog(
      "Export PDF", 
      `Génération du rapport de capabilité PDF professionnel pour le Master ${calcMasterId || "Manuel"}`
    );

    const styleElements: HTMLStyleElement[] = [];
    const originalStyles: string[] = [];
    const disabledSheets: StyleSheet[] = [];

    try {
      const element = document.getElementById("pdf-report-template");
      if (!element) {
        throw new Error("Modèle PDF introuvable dans le DOM.");
      }

      // 1. Sanitize style tags to bypass html2canvas oklch parsing crashes (Tailwind v4 compatibility)
      const styleEls = Array.from(document.querySelectorAll("style"));
      styleEls.forEach((el) => {
        styleElements.push(el);
        originalStyles.push(el.textContent || "");
        if (el.textContent && el.textContent.includes("oklch")) {
          let text = el.textContent;
          // Replace oklch values with grayscale rgb values based on lightness
          text = text.replace(/oklch\(\s*([0-9.%]+)\s+([0-9.%]+)\s+([0-9.%a-zA-Z_-]+)(?:\s*\/\s*([0-9.%a-zA-Z_()--]+))?\s*\)/g, (match, p1, p2, p3, p4) => {
            try {
              let lVal = parseFloat(p1);
              if (p1.includes('%')) lVal /= 100;
              const gray = Math.round(Math.max(0, Math.min(1, lVal)) * 255);
              if (p4) {
                if (p4.includes("var") || p4.includes("-")) {
                  return `rgb(${gray}, ${gray}, ${gray})`;
                }
                let alpha = parseFloat(p4);
                if (p4.includes('%')) alpha /= 100;
                return `rgba(${gray}, ${gray}, ${gray}, ${alpha})`;
              }
              return `rgb(${gray}, ${gray}, ${gray})`;
            } catch (e) {
              return "rgb(120, 120, 120)";
            }
          });
          // Fallback for any other oklch color strings
          text = text.replace(/oklch\([^)]+\)/g, "rgb(120, 120, 120)");
          el.textContent = text;
        }
      });

      // Disable any link/external stylesheets temporarily since they could also contain oklch
      for (let i = 0; i < document.styleSheets.length; i++) {
        const sheet = document.styleSheets[i];
        if (sheet.href) {
          try {
            sheet.disabled = true;
            disabledSheets.push(sheet);
          } catch (e) {
            // ignore
          }
        } else {
          try {
            const rules = sheet.cssRules || sheet.rules;
            if (rules) {
              let hasOklch = false;
              for (let j = 0; j < rules.length; j++) {
                if (rules[j].cssText.includes("oklch")) {
                  hasOklch = true;
                  break;
                }
              }
              if (hasOklch) {
                sheet.disabled = true;
                disabledSheets.push(sheet);
              }
            }
          } catch (e) {
            try {
              sheet.disabled = true;
              disabledSheets.push(sheet);
            } catch (e2) {
              // ignore
            }
          }
        }
      }

      // Temporarily show the template for rendering (placed off-screen absolute)
      element.style.display = "block";

      // Give a tiny timeout for dimensions/layout/fonts to fully stabilize
      await new Promise((resolve) => setTimeout(resolve, 300));

      const canvas = await html2canvas(element, {
        scale: 2.5, // Crisp high-definition vector text and charts
        useCORS: true,
        logging: false,
        backgroundColor: "#ffffff",
        allowTaint: true,
      });

      // Hide the template from DOM layout again
      element.style.display = "none";

      const imgData = canvas.toDataURL("image/jpeg", 0.98);
      const pdf = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4",
      });

      const pdfWidth = pdf.internal.pageSize.getWidth(); // 210mm
      const pdfHeight = pdf.internal.pageSize.getHeight(); // 297mm
      const imgWidth = canvas.width;
      const imgHeight = canvas.height;
      const ratio = imgWidth / imgHeight;
      const calcImgHeight = pdfWidth / ratio;

      // Add high-resolution image perfectly into the page boundaries
      pdf.addImage(imgData, "JPEG", 0, 0, pdfWidth, Math.min(calcImgHeight, pdfHeight));

      const filename = `Rapport_Capabilite_${calcMasterId || "MANUEL"}_${new Date().toISOString().slice(0, 10)}.pdf`;
      pdf.save(filename);

      showToast(`Rapport PDF '${filename}' exporté avec succès !`, "success");
    } catch (error: any) {
      console.error("Error generating PDF", error);
      showToast(`Échec de l'exportation du PDF : ${error?.message || "Erreur inconnue"}`, "error");
    } finally {
      // 2. Restore original styles and stylesheets
      styleElements.forEach((el, idx) => {
        el.textContent = originalStyles[idx];
      });
      disabledSheets.forEach((sheet) => {
        try {
          sheet.disabled = false;
        } catch (e) {
          // ignore
        }
      });
      setIsExportingPDF(false);
    }
  };

  // Save current capability calculation to history list
  const handleSaveCalculation = () => {
    if (!calcMasterId) {
      showToast("Veuillez d'abord lier ce calcul à un Master en cliquant sur 'Calculer' depuis la liste des Masters.", "error");
      return;
    }
    if (!activeStats || activeStats.n === 0) {
      showToast("Aucun résultat de mesure à enregistrer.", "error");
      return;
    }

    // Create calculation record
    const newCalc: CapabilityCalculation = {
      id: `CALC_${new Date().getTime()}`,
      title: `Contrôle Capabilité ${calcCharacteristic}`,
      date: new Date().toLocaleDateString('fr-FR'),
      testeur: calcTester,
      masterId: calcMasterId,
      characteristic: calcCharacteristic,
      nominal: calcNominal,
      lsl: calcLsl,
      usl: calcUsl,
      measurements: [...measurementsArray],
      operator: calcOperator || session?.name || "Technicien"
    };

    setCalculations(prev => {
      const updated = [...prev, newCalc];
      localStorage.setItem("mms_capability_calculations", JSON.stringify(updated));
      return updated;
    });

    // Record an entry in the logs
    addLog(
      "Enregistrement Capabilité", 
      `Nouveau calcul de capabilité enregistré pour le Master ${calcMasterId} (Cpk: ${activeStats.cpk.toFixed(3)}, Moyenne: ${activeStats.mean.toFixed(4)})`
    );

    showToast(`Calcul enregistré avec succès pour ${calcMasterId} !`, "success");
  };

  if (!session) {
    return (
      <div className="min-h-screen bg-slate-950 text-slate-200 font-mono selection:bg-sky-500 selection:text-slate-950 flex flex-col justify-center items-center relative overflow-hidden p-4 border-t-8 border-sky-500">
        
        {/* Background Grid Pattern (ACTIA style) */}
        <div className="absolute inset-0 bg-[linear-gradient(to_right,#0f172a_1px,transparent_1px),linear-gradient(to_bottom,#0f172a_1px,transparent_1px)] bg-[size:4rem_4rem] [mask-image:radial-gradient(ellipse_60%_50%_at_50%_50%,#000_70%,transparent_100%)] opacity-25 animate-pulse"></div>
        
        {/* Decorative background glows */}
        <div className="absolute -top-40 -left-40 w-96 h-96 bg-sky-500/10 rounded-full blur-3xl"></div>
        <div className="absolute -bottom-40 -right-40 w-96 h-96 bg-emerald-500/10 rounded-full blur-3xl"></div>

        {/* Toast Notification */}
        {toast && (
          <div className={`fixed top-5 right-5 z-50 flex items-center gap-3 px-5 py-3 rounded-none shadow-2xl border-2 bg-slate-900 ${toast.type === 'warning' ? 'border-amber-500' : 'border-sky-500'} max-w-sm font-mono animate-bounce`}>
            {toast.type === 'success' && <CheckCircle className="text-emerald-400 h-5 w-5 shrink-0" />}
            {toast.type === 'info' && <Info className="text-sky-400 h-5 w-5 shrink-0" />}
            {toast.type === 'error' && <XCircle className="text-rose-400 h-5 w-5 shrink-0" />}
            {toast.type === 'warning' && <AlertTriangle className="text-amber-400 h-5 w-5 shrink-0" />}
            <p className="text-xs font-bold text-slate-100">{toast.message}</p>
          </div>
        )}

        <div className="w-full max-w-lg bg-slate-900 border border-slate-800 p-8 shadow-2xl relative z-10">
          
          {/* Logo and Branding */}
          <div className="flex flex-col items-center text-center pb-6 border-b border-slate-800 mb-6">
            <ActiaLogo className="h-16 w-auto mb-4 hover:scale-105 transition-transform duration-300" />
            <h1 className="text-xl font-black uppercase tracking-wider text-white font-display">
              Master Management System
            </h1>
            <p className="text-[11px] text-slate-400 uppercase font-mono tracking-widest mt-1">
              Métrologie & Capabilité (FR 509)
            </p>
          </div>

          {/* Tab Selector for Session Type */}
          <div className="grid grid-cols-2 gap-2 p-1 bg-slate-950 border border-slate-800 mb-6">
            <button
              type="button"
              onClick={() => {
                setLoginRole('technician');
                setLoginMatricule("");
                setLoginName("");
              }}
              className={`py-3 text-xs uppercase font-bold tracking-wider flex items-center justify-center gap-2 transition-all cursor-pointer ${
                loginRole === 'technician'
                  ? 'bg-sky-500 text-slate-950 font-black'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <User className="h-4 w-4" />
              <span>Technicien</span>
            </button>
            <button
              type="button"
              onClick={() => {
                setLoginRole('admin');
                setLoginMatricule("");
                setLoginName("");
              }}
              className={`py-3 text-xs uppercase font-bold tracking-wider flex items-center justify-center gap-2 transition-all cursor-pointer ${
                loginRole === 'admin'
                  ? 'bg-amber-500 text-slate-950 font-black'
                  : 'text-slate-400 hover:text-white'
              }`}
            >
              <Shield className="h-4 w-4" />
              <span>Administrateur</span>
            </button>
          </div>

          {/* Login Form */}
          <form onSubmit={handleLogin} className="space-y-4">
            
            {/* Selection Dropdown for Technician */}
            {loginRole === 'technician' && (
              <div className="animate-fadeIn">
                <label className="block text-[10px] font-bold text-sky-400 uppercase tracking-widest mb-1.5">
                  Sélectionner un Technicien Enregistré *
                </label>
                <select
                  onChange={(e) => {
                    const selectedMat = e.target.value;
                    if (selectedMat) {
                      const found = technicians.find(t => t.matricule === selectedMat);
                      if (found) {
                        setLoginMatricule(found.matricule);
                        setLoginName(found.name);
                      }
                    } else {
                      setLoginMatricule("");
                      setLoginName("");
                    }
                  }}
                  className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2.5 text-sm text-slate-100 font-mono focus:outline-none focus:border-sky-500 focus:ring-1 focus:ring-sky-500"
                >
                  <option value="">-- Choisissez un Technicien --</option>
                  {technicians.map((t) => (
                    <option key={t.matricule} value={t.matricule}>
                      {t.name} ({t.matricule})
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Common Matricule Input */}
            <div>
              <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                {loginRole === 'admin' ? "Matricule Admin *" : "Matricule Technicien *"}
              </label>
              <div className="relative">
                <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-500">
                  <Key className="h-4 w-4" />
                </span>
                <input
                  type="text"
                  placeholder={loginRole === 'admin' ? "Ex: ADM01" : "Ex: MTR9823"}
                  value={loginMatricule}
                  onChange={(e) => setLoginMatricule(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-none pl-10 pr-4 py-2.5 text-sm text-slate-100 placeholder-slate-600 font-mono focus:outline-none focus:border-sky-500 focus:ring-1 focus:ring-sky-500"
                  required
                />
              </div>
            </div>

            {/* Name Input - Required for Tech, Optional for Admin */}
            <div>
              <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                {loginRole === 'admin' ? "Nom de l'Administrateur (Optionnel)" : "Nom du Technicien *"}
              </label>
              <div className="relative">
                <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-500">
                  <User className="h-4 w-4" />
                </span>
                <input
                  type="text"
                  placeholder="Ex: BEN MANSOUR Samir"
                  value={loginName}
                  onChange={(e) => setLoginName(e.target.value)}
                  className="w-full bg-slate-950 border border-slate-800 rounded-none pl-10 pr-4 py-2.5 text-sm text-slate-100 placeholder-slate-600 font-mono focus:outline-none focus:border-sky-500 focus:ring-1 focus:ring-sky-500"
                  required={loginRole === 'technician'}
                />
              </div>
            </div>

            {/* Admin Code Password Input (Only for Admin) */}
            {loginRole === 'admin' && (
              <div className="animate-fadeIn">
                <div className="flex justify-between mb-1.5">
                  <label className="block text-[10px] font-bold text-amber-400 uppercase tracking-widest">
                    Code d'accès Admin *
                  </label>
                </div>
                <div className="relative">
                  <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-amber-500/70">
                    <Lock className="h-4 w-4" />
                  </span>
                  <input
                    type="password"
                    placeholder="Saisir le code d'accès"
                    value={loginAdminCode}
                    onChange={(e) => setLoginAdminCode(e.target.value)}
                    className="w-full bg-slate-950 border border-amber-900/30 rounded-none pl-10 pr-4 py-2.5 text-sm text-slate-100 placeholder-slate-600 font-mono focus:outline-none focus:border-amber-500 focus:ring-1 focus:ring-amber-500"
                    required
                  />
                </div>
              </div>
            )}

            {/* Submit Button */}
            <button
              type="submit"
              className={`w-full py-3.5 px-4 font-black text-xs uppercase tracking-widest mt-6 cursor-pointer border shadow-lg transition-all flex items-center justify-center gap-2 ${
                loginRole === 'admin'
                  ? 'bg-amber-500 border-amber-400 text-slate-950 hover:bg-amber-400 shadow-amber-500/5'
                  : 'bg-sky-500 border-sky-400 text-slate-950 hover:bg-sky-400 shadow-sky-500/5'
              }`}
            >
              <LogIn className="h-4 w-4" />
              <span>Démarrer la Session</span>
            </button>
          </form>

          {/* Footer inside login block */}
          <div className="mt-8 pt-4 border-t border-slate-800/50 flex justify-between text-[9px] text-slate-500 font-mono uppercase">
            <span>BANC DE CONTROLE FR 509</span>
            <span>ACTIA AUTOMOTIVE</span>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-950 text-slate-200 font-sans selection:bg-sky-500 selection:text-slate-950 pb-16 border-t-8 border-sky-500">
      
      {/* Toast Notification Header */}
      {toast && (
        <div className={`fixed top-5 right-5 z-50 flex items-center gap-3 px-5 py-3 rounded-none shadow-2xl border-2 bg-slate-900 ${toast.type === 'warning' ? 'border-amber-500' : 'border-sky-500'} max-w-sm animate-bounce font-mono`}>
          {toast.type === 'success' && <CheckCircle className="text-emerald-400 h-5 w-5 shrink-0" />}
          {toast.type === 'info' && <Info className="text-sky-400 h-5 w-5 shrink-0" />}
          {toast.type === 'error' && <XCircle className="text-rose-400 h-5 w-5 shrink-0" />}
          {toast.type === 'warning' && <AlertTriangle className="text-amber-400 h-5 w-5 shrink-0" />}
          <p className="text-xs font-bold text-slate-100">{toast.message}</p>
        </div>
      )}

      {/* Main Container Header - Geometric Balance Styling */}
      <header className="border-b border-slate-900 bg-slate-900/90 backdrop-blur sticky top-0 z-40 h-20 flex items-center">
        <div className="max-w-7xl mx-auto w-full px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between gap-4">
            
            {/* Logo and Brand Title */}
            <div className="flex items-center gap-3">
              <ActiaLogo className="h-11 w-auto shrink-0 hover:scale-[1.02] transition-transform duration-200" />
              <div className="hidden xs:block h-8 w-px bg-slate-800 self-center mx-1"></div>
              <div>
                <h1 className="text-base md:text-lg font-black tracking-tight uppercase text-white font-display flex items-center gap-2">
                  Master Management System
                  <span className="text-[10px] font-bold bg-sky-500/10 text-sky-400 border border-sky-500/20 px-2 py-0.5 rounded-none font-mono">
                    v2.1
                  </span>
                </h1>
                <p className="text-[11px] text-slate-400 uppercase font-mono tracking-wider">Métrologie, Capabilité & Suivi (FR 509)</p>
              </div>
            </div>

            {/* Industrial Metadata & Connection status */}
            <div className="flex items-center gap-4 sm:gap-6 text-[10px] font-mono uppercase text-slate-400">
              <div className="hidden lg:flex flex-col text-right border-r border-slate-800 pr-6">
                <div>{session?.role === 'admin' ? "Administrateur" : "Technicien"} : <span className="text-sky-400 font-bold">{session?.name}</span></div>
                <div>Matricule : <span className="text-slate-300 font-bold">{session?.matricule}</span></div>
              </div>
              <div className="hidden sm:flex flex-col text-right">
                <div>Session : <span className={`font-bold ${session?.role === 'admin' ? 'text-amber-400' : 'text-sky-400'}`}>{session?.role === 'admin' ? 'ADMINISTRATEUR' : 'TECHNICIEN'}</span></div>
                <div>Banc : <span className="text-emerald-400 font-bold">FR 509</span></div>
              </div>
              
              {/* Logout Button */}
              <button
                onClick={handleLogout}
                className="p-2 bg-slate-900 hover:bg-rose-950/40 border border-slate-800 hover:border-rose-900/60 text-slate-400 hover:text-rose-400 transition-all flex items-center gap-1.5 cursor-pointer h-8"
                title="Fermer la session"
              >
                <LogOut className="h-3.5 w-3.5" />
                <span className="hidden md:inline text-[9px] font-bold">Quitter</span>
              </button>

              <div className="w-8 h-8 rounded-full bg-slate-900 border border-slate-800 flex items-center justify-center shrink-0" title="Base de données connectée localement">
                <div className="w-2.5 h-2.5 rounded-full bg-emerald-500 shadow-[0_0_10px_#10b981] animate-pulse"></div>
              </div>
            </div>

          </div>
        </div>
      </header>

      {/* Main Content Workspace */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 pt-8">
        
        {/* Navigation Tabs - Geometric Balance visual pattern with 3 or 4 prominent buttons */}
        <div className={`grid grid-cols-1 ${session?.role === 'admin' ? 'lg:grid-cols-4 md:grid-cols-2' : 'md:grid-cols-3'} gap-4 mb-8`}>
          
          {/* BOUTON 1: Circle Focus shape (Testeurs) */}
          <button
            id="btn_tab_testeurs"
            onClick={() => setActiveTab('testeurs')}
            className={`group relative p-6 border-2 transition-all duration-300 text-left overflow-hidden flex items-center justify-between ${
              activeTab === 'testeurs'
                ? 'bg-slate-900 border-sky-500 shadow-xl shadow-sky-500/5'
                : 'bg-slate-900/40 border-slate-900 hover:border-slate-800 hover:bg-slate-900/60'
            }`}
          >
            {/* Corner tag indicator */}
            <div className="absolute top-2 right-3 text-[9px] text-slate-500 font-mono font-bold tracking-widest">01 / ROUND</div>
            
            <div className="flex items-center gap-4">
              {/* Circular focus container */}
              <div className={`w-14 h-14 rounded-full border-4 flex items-center justify-center transition-all duration-300 shrink-0 ${
                activeTab === 'testeurs'
                  ? 'border-sky-500 bg-sky-950/50 text-sky-400'
                  : 'border-slate-800 bg-slate-900 text-slate-400 group-hover:border-slate-700'
              }`}>
                <Layers className="h-6 w-6" />
              </div>
              <div>
                <div className="text-[10px] text-sky-400 font-mono uppercase tracking-wider font-bold">BOUTON 1</div>
                <div className="font-display font-bold text-sm md:text-base text-white tracking-tight uppercase">Visualiser Testeurs</div>
                <p className="text-[11px] text-slate-500 font-mono mt-0.5">{masters.length} Masters enregistrés</p>
              </div>
            </div>
            
            {/* Tiny arrow */}
            <ArrowRight className={`h-4 w-4 transition-transform duration-300 ${activeTab === 'testeurs' ? 'text-sky-400 translate-x-1' : 'text-slate-700'}`} />
          </button>

          {/* BOUTON 2: Sharp Square Focus shape (Créer new master) */}
          <button
            id="btn_tab_nouveau"
            onClick={() => {
              setActiveTab('nouveau-master');
            }}
            className={`group relative p-6 border-2 transition-all duration-300 text-left overflow-hidden flex items-center justify-between ${
              activeTab === 'nouveau-master'
                ? 'bg-slate-900 border-emerald-500 shadow-xl shadow-emerald-500/5'
                : 'bg-slate-900/40 border-slate-900 hover:border-slate-800 hover:bg-slate-900/60'
            }`}
          >
            {/* Corner tag indicator */}
            <div className="absolute top-2 right-3 text-[9px] text-slate-500 font-mono font-bold tracking-widest">
              02 / SQUARE
            </div>
            
            <div className="flex items-center gap-4">
              {/* Square focus container */}
              <div className={`w-14 h-14 rounded-none border-4 flex items-center justify-center transition-all duration-300 shrink-0 ${
                activeTab === 'nouveau-master'
                  ? 'border-emerald-500 bg-emerald-950/50 text-emerald-400'
                  : 'border-slate-800 bg-slate-900 text-slate-400 group-hover:border-slate-700'
              }`}>
                <PlusCircle className="h-6 w-6" />
              </div>
              <div>
                <div className="text-[10px] text-emerald-400 font-mono uppercase tracking-wider font-bold">
                  BOUTON 2
                </div>
                <div className="font-display font-bold text-sm md:text-base text-white tracking-tight uppercase flex items-center gap-1.5">
                  Créer New Master
                </div>
                <p className="text-[11px] text-slate-500 font-mono mt-0.5">
                  Enregistrer pièce étalon
                </p>
              </div>
            </div>
            
            {/* Tiny arrow */}
            <ArrowRight className={`h-4 w-4 transition-transform duration-300 ${activeTab === 'nouveau-master' ? 'text-emerald-400 translate-x-1' : 'text-slate-700'}`} />
          </button>

          {/* BOUTON 3: Diamond Focus shape (Calcul capabilité) */}
          <button
            id="btn_tab_calcul"
            onClick={() => setActiveTab('calcul')}
            className={`group relative p-6 border-2 transition-all duration-300 text-left overflow-hidden flex items-center justify-between ${
              activeTab === 'calcul'
                ? 'bg-slate-900 border-rose-500 shadow-xl shadow-rose-500/5'
                : 'bg-slate-900/40 border-slate-900 hover:border-slate-800 hover:bg-slate-900/60'
            }`}
          >
            {/* Corner tag indicator */}
            <div className="absolute top-2 right-3 text-[9px] text-slate-500 font-mono font-bold tracking-widest">03 / DIAMOND</div>
            
            <div className="flex items-center gap-4">
              {/* Diamond focus container */}
              <div className="w-14 h-14 flex items-center justify-center shrink-0">
                <div className={`w-10 h-10 border-2 rotate-45 flex items-center justify-center transition-all duration-300 ${
                  activeTab === 'calcul'
                    ? 'border-rose-500 bg-rose-950/50 text-rose-400'
                    : 'border-slate-800 bg-slate-900 text-slate-400 group-hover:border-slate-700'
                }`}>
                  <Calculator className="h-5 w-5 -rotate-45" />
                </div>
              </div>
              <div>
                <div className="text-[10px] text-rose-400 font-mono uppercase tracking-wider font-bold">BOUTON 3</div>
                <div className="font-display font-bold text-sm md:text-base text-white tracking-tight uppercase">Calcul CPK & MSA</div>
                <p className="text-[11px] text-slate-500 font-mono mt-0.5">Cpk, Sigma, Répétabilité</p>
              </div>
            </div>
            
            {/* Tiny status blinker */}
            <span className="flex h-2 w-2 relative">
              <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span>
              <span className="relative inline-flex rounded-full h-2 w-2 bg-rose-500"></span>
            </span>
          </button>

          {/* BOUTON 4: Shield Focus shape (Admin Panel: Techniciens & Historique) */}
          {session?.role === 'admin' && (
            <button
              id="btn_tab_admin_panel"
              onClick={() => setActiveTab('admin-panel')}
              className={`group relative p-6 border-2 transition-all duration-300 text-left overflow-hidden flex items-center justify-between ${
                activeTab === 'admin-panel'
                  ? 'bg-slate-900 border-amber-500 shadow-xl shadow-amber-500/5'
                  : 'bg-slate-900/40 border-slate-900 hover:border-slate-800 hover:bg-slate-900/60'
              }`}
            >
              {/* Corner tag indicator */}
              <div className="absolute top-2 right-3 text-[9px] text-slate-500 font-mono font-bold tracking-widest">04 / SHIELD</div>
              
              <div className="flex items-center gap-4">
                {/* Shield focus container */}
                <div className={`w-14 h-14 rounded-full border-4 flex items-center justify-center transition-all duration-300 shrink-0 ${
                  activeTab === 'admin-panel'
                    ? 'border-amber-500 bg-amber-950/50 text-amber-400'
                    : 'border-slate-800 bg-slate-900 text-slate-400 group-hover:border-slate-700'
                }`}>
                  <Shield className="h-6 w-6" />
                </div>
                <div>
                  <div className="text-[10px] text-amber-400 font-mono uppercase tracking-wider font-bold font-mono">ADMIN ONLY</div>
                  <div className="font-display font-bold text-sm md:text-base text-white tracking-tight uppercase">Admin & Historique</div>
                  <p className="text-[11px] text-slate-500 font-mono mt-0.5">Comptes & Logs ({logs.length})</p>
                </div>
              </div>
              
              {/* Tiny arrow */}
              <ArrowRight className={`h-4 w-4 transition-transform duration-300 ${activeTab === 'admin-panel' ? 'text-amber-400 translate-x-1' : 'text-slate-700'}`} />
            </button>
          )}
        </div>


        {/* ========================================================================= */}
        {/* TAB 1: LISTE DES TESTEURS & PIECES MASTERS                                */}
        {/* ========================================================================= */}
        {activeTab === 'testeurs' && (
          <div className="space-y-6 animate-fadeIn">
            
            {/* Realtime KPI stats panel */}
            <div className="grid grid-cols-2 lg:grid-cols-5 gap-4">
              <div className="bg-slate-900/60 border border-slate-900 border-l-4 border-l-sky-500 p-4 rounded-none space-y-1 shadow-md">
                <span className="text-[10px] text-slate-400 uppercase tracking-widest font-mono">Total Masters</span>
                <div className="text-3xl font-black font-display text-white">{masters.length}</div>
                <div className="text-[10px] text-slate-500 uppercase font-mono">Tous testeurs</div>
              </div>

              <div className="bg-slate-900/60 border border-slate-900 border-l-4 border-l-emerald-500 p-4 rounded-none space-y-1 shadow-md">
                <span className="text-[10px] text-emerald-400 uppercase tracking-widest font-mono flex items-center gap-1.5 font-bold">
                  <span className="h-2 w-2 rounded-full bg-emerald-500 inline-block animate-pulse shadow-[0_0_8px_#10b981]"></span>
                  Actives
                </span>
                <div className="text-3xl font-black font-display text-emerald-400">
                  {masters.filter(m => normalizeStatus(m.statut) === "active").length}
                </div>
                <div className="text-[10px] text-slate-500 uppercase font-mono">En production</div>
              </div>

              <div className="bg-slate-900/60 border border-slate-900 border-l-4 border-l-amber-500 p-4 rounded-none space-y-1 shadow-md">
                <span className="text-[10px] text-amber-400 uppercase tracking-widest font-mono font-bold">Obsolètes</span>
                <div className="text-3xl font-black font-display text-amber-400">
                  {masters.filter(m => normalizeStatus(m.statut) === "obsolete").length}
                </div>
                <div className="text-[10px] text-slate-500 uppercase font-mono">Fin de vie / Redesign</div>
              </div>

              <div className="bg-slate-900/60 border border-slate-900 border-l-4 border-l-rose-500 p-4 rounded-none space-y-1 shadow-md">
                <span className="text-[10px] text-rose-400 uppercase tracking-widest font-mono font-bold">Endommagées</span>
                <div className="text-3xl font-black font-display text-rose-400">
                  {masters.filter(m => normalizeStatus(m.statut) === "endommagee").length}
                </div>
                <div className="text-[10px] text-slate-500 uppercase font-mono">À remplacer</div>
              </div>

              <div className="bg-slate-900/60 border border-slate-900 border-l-4 border-l-sky-400 p-4 rounded-none space-y-1 col-span-2 lg:col-span-1 shadow-md">
                <span className="text-[10px] text-sky-400 uppercase tracking-widest font-mono font-bold">Vérifiées OK</span>
                <div className="text-3xl font-black font-display text-sky-300">
                  {masters.filter(m => m.verif === "OK").length}
                </div>
                <div className="text-[10px] text-slate-500 uppercase font-mono">Conformité métrologie</div>
              </div>
            </div>

            {/* SPC Trend Warning Banner */}
            {driftingMastersCount > 0 && (
              <div className="bg-amber-500/10 border border-amber-500/30 p-4 rounded-none flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 font-mono text-xs animate-fadeIn">
                <div className="flex items-start gap-3">
                  <AlertTriangle className="h-5 w-5 text-amber-500 shrink-0 mt-0.5" />
                  <div>
                    <h5 className="font-bold text-amber-400 uppercase tracking-wider text-xs flex items-center gap-1.5">
                      <span className="h-1.5 w-1.5 rounded-full bg-amber-500 inline-block animate-pulse"></span>
                      ALERTE SPC : DÉRIVE DE CAPABILITÉ DÉTECTÉE (&gt; 5%)
                    </h5>
                    <p className="text-slate-300 mt-1 font-sans text-[11px] leading-relaxed">
                      {driftingMastersCount} master(s) présente(nt) une dérive statistique de plus de 5% de la bande de tolérance vers les limites (LSL/USL) sur la moyenne des 3 derniers calculs de capabilité.
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => {
                    const driftingIds = masters.filter(m => {
                      const trend = getMasterTrend(m.idMaster);
                      return trend && trend.isWarning;
                    }).map(m => m.idMaster);
                    
                    if (driftingIds.length > 0) {
                      setSearchQueryInput(driftingIds[0]);
                      setSearchQuery(driftingIds[0]);
                      showToast(`Filtré sur le premier master en dérive : ${driftingIds[0]}`, "info");
                    }
                  }}
                  className="bg-amber-500 hover:bg-amber-400 text-slate-950 font-black px-3.5 py-2 rounded-none uppercase tracking-wider text-[9px] whitespace-nowrap transition-all self-start sm:self-center cursor-pointer border border-amber-400"
                >
                  Identifier les Masters en Dérive
                </button>
              </div>
            )}

            {/* Filters dashboard - Geometric Balance Style */}
            <div className="bg-slate-900/40 border border-slate-900 p-6 rounded-none space-y-4 shadow-md">
              <div className="flex items-center justify-between border-b border-slate-900 pb-3">
                <div className="flex items-center gap-2 text-slate-300 font-bold text-xs uppercase tracking-wider font-display">
                  <Filter className="h-4 w-4 text-sky-500" />
                  <span>Filtres de recherche rapide</span>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                {/* Search query input */}
                <div>
                  <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono mb-1.5">01 // Rechercher</label>
                  <div className="relative">
                    <Search className="absolute left-3 top-2.5 h-4 w-4 text-slate-600" />
                    <input
                      id="search_masters"
                      type="text"
                      placeholder="ID, Réf, N° Série..."
                      value={searchQueryInput}
                      onChange={(e) => setSearchQueryInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          handleApplyFilters();
                        }
                      }}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none pl-9 pr-4 py-2 text-xs text-slate-100 placeholder-slate-600 focus:outline-none focus:border-sky-500 font-mono"
                    />
                  </div>
                </div>

                {/* Tester selection */}
                <div>
                  <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono mb-1.5">02 // Banc de Destination</label>
                  <select
                    id="filter_tester"
                    value={selectedTesterFilterInput}
                    onChange={(e) => {
                      const val = e.target.value;
                      setSelectedTesterFilterInput(val);
                      if (val !== "All") {
                        setStatusFilterInput("All");
                      }
                    }}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500 font-mono"
                  >
                    <option value="All">Tous les Testeurs ({uniqueTesters.length})</option>
                    {uniqueTesters.map(tester => (
                      <option key={tester} value={tester}>{tester}</option>
                    ))}
                  </select>
                </div>

                {/* Status Selection */}
                <div>
                  <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono mb-1.5">03 // Statut Pièce</label>
                  <select
                    id="filter_status"
                    value={statusFilterInput}
                    onChange={(e) => setStatusFilterInput(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500 font-mono"
                  >
                    <option value="All">Tous les Statuts</option>
                    <option value="Active">Active</option>
                    <option value="Obsolète">Obsolète</option>
                    <option value="Endommagée">Endommagée</option>
                  </select>
                </div>

                {/* Verification Selection */}
                <div>
                  <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono mb-1.5">04 // Métrologie</label>
                  <select
                    id="filter_verif"
                    value={verifFilterInput}
                    onChange={(e) => setVerifFilterInput(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500 font-mono"
                  >
                    <option value="All">Toutes les Vérif</option>
                    <option value="OK">OK</option>
                    <option value="KO">KO</option>
                    <option value="FIN DE VIE">FIN DE VIE</option>
                  </select>
                </div>
              </div>

              {/* Validation Action and Indicator */}
              <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pt-4 border-t border-slate-900/60">
                <div className="flex items-center gap-2">
                  {hasPendingFilterChanges ? (
                    <div className="flex items-center gap-2 text-amber-400 font-mono text-[10px] uppercase animate-pulse">
                      <span className="h-2.5 w-2.5 rounded-full bg-amber-500 inline-block"></span>
                      <span>Modifications en attente de validation</span>
                    </div>
                  ) : (
                    <div className="flex items-center gap-2 text-slate-500 font-mono text-[10px] uppercase">
                      <span className="h-2 w-2 rounded-full bg-slate-800 inline-block"></span>
                      <span>Filtres synchronisés avec la liste</span>
                    </div>
                  )}
                </div>

                <div className="flex flex-col sm:flex-row gap-2 w-full sm:w-auto">
                  <button
                    id="btn_reset_filters_bottom"
                    type="button"
                    onClick={() => {
                      setSearchQuery("");
                      setSelectedTesterFilter("All");
                      setStatusFilter("All");
                      setVerifFilter("All");
                      setSearchQueryInput("");
                      setSelectedTesterFilterInput("All");
                      setStatusFilterInput("All");
                      setVerifFilterInput("All");
                      showToast("Filtres réinitialisés.", "info");
                    }}
                    className="w-full sm:w-auto text-xs px-6 py-2.5 font-mono uppercase tracking-wider font-bold transition-all flex items-center justify-center gap-2 border border-slate-800 bg-slate-950 hover:bg-slate-900 text-slate-400 hover:text-slate-200 cursor-pointer"
                  >
                    <RefreshCw className="h-4 w-4 shrink-0" />
                    <span>Réinitialiser</span>
                  </button>

                  <button
                    id="btn_apply_filters"
                    type="button"
                    onClick={handleApplyFilters}
                    className={`w-full sm:w-auto text-xs px-6 py-2.5 font-mono uppercase tracking-wider font-bold transition-all flex items-center justify-center gap-2 border cursor-pointer ${
                      hasPendingFilterChanges
                        ? "bg-sky-500 hover:bg-sky-400 text-slate-950 border-sky-400 font-black shadow-[0_0_15px_rgba(14,165,233,0.25)]"
                        : "bg-slate-950 hover:bg-slate-900 text-slate-400 border-slate-800"
                    }`}
                  >
                    <Check className="h-4 w-4 shrink-0" />
                    <span>Valider la recherche / Appliquer</span>
                  </button>
                </div>
              </div>

              {/* Database Import Section */}
              {true && (
                <div className="pt-2 border-t border-slate-900 mt-2 space-y-3">
                  <div className="flex flex-col sm:flex-row gap-3">
                    <button
                      id="btn_toggle_excel_importer"
                      onClick={() => setShowCsvImporter(!showCsvImporter)}
                      className="text-[10px] bg-slate-950 hover:bg-slate-900 text-slate-100 border border-slate-800 hover:border-slate-700 rounded-none px-4 py-2.5 font-mono uppercase tracking-wider flex items-center gap-2 transition-all font-bold cursor-pointer"
                    >
                      <FileSpreadsheet className="h-4 w-4 text-emerald-400" />
                      <span>{showCsvImporter ? "Masquer l'importateur de Fichiers" : "Mise à jour de la base (Fichier Excel / CSV)"}</span>
                    </button>
                  </div>

                  {showCsvImporter && (
                    <div className="p-5 bg-slate-950/80 border border-slate-900 space-y-4 animate-fadeIn">
                      
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        {/* Drag and Drop Zone */}
                        <div className="space-y-2">
                          <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono">Option A : Charger un fichier Excel (.xlsx, .xls) ou CSV (.csv, .txt)</label>
                          <div
                            onDragOver={(e) => {
                              e.preventDefault();
                              setIsDraggingExcel(true);
                            }}
                            onDragLeave={() => setIsDraggingExcel(false)}
                            onDrop={(e) => {
                              e.preventDefault();
                              setIsDraggingExcel(false);
                              const file = e.dataTransfer.files?.[0];
                              if (file) handleExcelFileParse(file);
                            }}
                            onClick={() => document.getElementById("db_file_input")?.click()}
                            className={`border-2 border-dashed p-6 text-center cursor-pointer transition-all h-[130px] flex flex-col items-center justify-center ${
                              isDraggingExcel 
                                ? "border-emerald-500 bg-emerald-950/20 text-emerald-400" 
                                : "border-slate-800 hover:border-slate-700 bg-slate-950/40 text-slate-400"
                            }`}
                          >
                            <input
                              id="db_file_input"
                              type="file"
                              accept=".xlsx,.xls,.csv,.txt"
                              className="hidden"
                              onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (file) handleExcelFileParse(file);
                              }}
                            />
                            <Upload className="h-6 w-6 text-emerald-500 mb-2 animate-bounce" />
                            <p className="text-[10px] font-bold uppercase tracking-wider text-slate-200">Glisser-déposer le fichier de suivi</p>
                            <p className="text-[9px] text-slate-500 mt-1 font-sans">Supporte .xlsx, .xls, .csv ou .txt</p>
                          </div>
                        </div>

                        {/* Text paste zone */}
                        <div className="space-y-2">
                          <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono">Option B : Collage direct de lignes semi-colonnes (;)</label>
                          <textarea
                            id="raw_csv_textarea"
                            placeholder="UFT SMOOVE RTS;MASTER_452;SY4360;50DE6F;SY4360;;19/12/2026;;Active;;OK"
                            value={rawCsvPaste}
                            onChange={(e) => setRawCsvPaste(e.target.value)}
                            rows={4}
                            className="w-full bg-slate-900 border border-slate-800 rounded-none p-2.5 font-mono text-[11px] text-slate-200 placeholder-slate-700 focus:outline-none focus:border-emerald-500 h-[130px] resize-none"
                          />
                        </div>
                      </div>

                      {/* Submit for text paste option */}
                      {rawCsvPaste.trim() && (
                        <div className="flex justify-end border-t border-slate-900 pt-3">
                          <button
                            id="btn_csv_import_submit"
                            onClick={() => {
                              const items = parseUploadedCsvContent(rawCsvPaste);
                              if (items.length > 0) {
                                setExcelImportDetails({
                                  fileName: "Texte collé manuellement",
                                  rowCount: items.length,
                                  parsedItems: items
                                });
                                setRawCsvPaste("");
                              } else {
                                showToast("Aucune donnée valide n'a été détectée. Vérifiez le format.", "error");
                              }
                            }}
                            className="bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold px-4 py-2 text-xs rounded-none uppercase font-mono tracking-wider transition-colors"
                          >
                            Analyser le texte collé
                          </button>
                        </div>
                      )}

                      {/* Excel/CSV parsed preview and actions */}
                      {excelImportDetails && (
                        <div className="bg-slate-950 border border-emerald-900/50 p-4 space-y-4 font-mono animate-fadeIn">
                          <div className="flex items-start justify-between border-b border-slate-900 pb-2">
                            <div className="space-y-1">
                              <span className="text-[9px] text-slate-500 block uppercase">Fichier analysé</span>
                              <span className="text-xs font-bold text-emerald-400 block break-all">{excelImportDetails.fileName}</span>
                            </div>
                            <button
                              type="button"
                              onClick={() => setExcelImportDetails(null)}
                              className="text-slate-500 hover:text-rose-400 text-xs font-bold"
                            >
                              ✕ Annuler
                            </button>
                          </div>

                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs text-slate-400">
                            <div>Pièces masters détectées : <strong className="text-white text-sm">{excelImportDetails.rowCount}</strong></div>
                            <div className="text-[11px] text-slate-500 leading-normal">
                              Format détecté : colonnes de testeur, ID, référence de produit, série de produit, etc.
                            </div>
                          </div>

                          {/* Preview of first 3 rows in a mini-table */}
                          <div className="bg-slate-900/60 border border-slate-800 p-2.5 max-h-36 overflow-y-auto select-none rounded-none">
                            <span className="text-[9px] text-slate-500 uppercase font-bold block mb-1">Aperçu des 3 premières lignes :</span>
                            <div className="space-y-2">
                              {excelImportDetails.parsedItems.slice(0, 3).map((item, idx) => (
                                <div key={idx} className="text-[10px] text-slate-300 border-b border-slate-900/50 pb-1.5 last:border-0 last:pb-0">
                                  <span className="text-emerald-400 font-bold">{item.idMaster}</span> - {item.testeur} ({item.refProduitMaster || "SANS REF"}) 
                                  <span className="text-slate-500 ml-2">[{item.statut}] [{item.verif}]</span>
                                </div>
                              ))}
                              {excelImportDetails.parsedItems.length > 3 && (
                                <div className="text-[9px] text-slate-500 text-center italic pt-1">
                                  ... et {excelImportDetails.parsedItems.length - 3} autres lignes ...
                                </div>
                              )}
                            </div>
                          </div>

                          <div className="flex flex-col sm:flex-row gap-3 pt-2">
                            {/* OPTION 1: Overwrite */}
                            <button
                              type="button"
                              onClick={() => handleApplyDatabaseImport('replace')}
                              className="flex-1 bg-rose-500 hover:bg-rose-400 text-white font-bold py-2 px-4 uppercase text-center text-xs tracking-wider transition-colors cursor-pointer border border-rose-600"
                            >
                              ⚠️ Remplacer toute la base (Écraser)
                            </button>
                            
                            {/* OPTION 2: Merge */}
                            <button
                              type="button"
                              onClick={() => handleApplyDatabaseImport('merge')}
                              className="flex-1 bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold py-2 px-4 uppercase text-center text-xs tracking-wider transition-colors cursor-pointer"
                            >
                              🔄 Synchroniser & Fusionner (Mise à jour)
                            </button>
                          </div>

                          <p className="text-[10px] text-slate-500 leading-normal font-sans italic text-center">
                            Conseil : Choisissez "Remplacer" pour réinitialiser la base à l'état exact de votre fichier Excel de suivi. Choisissez "Synchroniser" pour ajouter ou modifier les données sans supprimer vos autres saisies manuelles.
                          </p>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Masters database table list */}
            {/* Masters database table list - Geometric Balance Style */}
            <div className="bg-slate-900/60 border border-slate-900 rounded-none overflow-hidden shadow-xl">
              <div className="px-6 py-5 border-b border-slate-900 bg-slate-900/40 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                <div>
                  <h3 className="font-black text-white text-base font-display uppercase tracking-tight">Catalogue des Pièces Masters</h3>
                  <p className="text-xs text-slate-500 font-mono">Affichage de {filteredMasters.length} sur {masters.length} masters au total</p>
                </div>
                
                {filteredMasters.length > 0 && (
                  <div className="relative">
                    <button
                      id="btn_export_masters"
                      onClick={() => setShowExportMenu(prev => !prev)}
                      className="text-[10px] text-slate-300 bg-slate-950 hover:bg-slate-900 border border-slate-800 rounded-none px-4 py-2 font-mono uppercase font-bold tracking-wider flex items-center gap-2 transition-all"
                    >
                      <Download className="h-3.5 w-3.5 text-sky-500" />
                      Exporter la sélection
                      <ChevronDown className="h-3.5 w-3.5 text-slate-500" />
                    </button>

                    {showExportMenu && (
                      <div className="absolute right-0 mt-1 z-20 bg-slate-950 border border-slate-800 shadow-xl min-w-[200px]">
                        <button
                          onClick={() => {
                            setShowExportMenu(false);
                            exportMastersCSV(filteredMasters);
                          }}
                          className="w-full text-left text-[10px] text-slate-300 hover:bg-slate-900 px-4 py-2 font-mono uppercase font-bold tracking-wider flex items-center gap-2 transition-all"
                        >
                          <FileText className="h-3.5 w-3.5 text-emerald-400" />
                          Exporter en CSV
                        </button>
                        <button
                          onClick={() => {
                            setShowExportMenu(false);
                            exportMastersXLSX(filteredMasters);
                          }}
                          className="w-full text-left text-[10px] text-slate-300 hover:bg-slate-900 px-4 py-2 font-mono uppercase font-bold tracking-wider flex items-center gap-2 transition-all border-t border-slate-900"
                        >
                          <FileSpreadsheet className="h-3.5 w-3.5 text-sky-400" />
                          Exporter en Excel (.xlsx)
                        </button>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {filteredMasters.length === 0 ? (
                <div className="p-16 text-center space-y-4">
                  <AlertTriangle className="h-10 w-10 text-slate-600 mx-auto" />
                  <p className="text-slate-300 font-bold font-display uppercase tracking-wider text-xs">Aucune pièce master correspondante</p>
                  <p className="text-xs text-slate-500 font-mono">Essayez de réinitialiser vos filtres ou de créer un nouveau master.</p>
                  <button
                    onClick={() => {
                      setSearchQuery("");
                      setSelectedTesterFilter("All");
                      setStatusFilter("All");
                      setVerifFilter("All");
                    }}
                    className="bg-sky-500 hover:bg-sky-400 text-slate-950 font-bold font-mono uppercase tracking-wider text-[10px] px-4 py-2 rounded-none transition-all"
                  >
                    Effacer les filtres
                  </button>
                </div>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-slate-950 border-b border-slate-900 text-slate-400 font-mono text-[9px] uppercase tracking-widest font-bold">
                        <th className="py-4 px-6">Identification Master</th>
                        <th className="py-4 px-6">Testeur de destination</th>
                        <th className="py-4 px-6">Réf Produit & N° Série</th>
                        <th className="py-4 px-6">Réf Carte & N° Série</th>
                        <th className="py-4 px-6">Date création</th>
                        <th className="py-4 px-6">Statut</th>
                        <th className="py-4 px-6">Vérif</th>
                        <th className="py-4 px-6 text-right">Actions rapides</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-900/60 font-mono">
                      {filteredMasters.map((item, index) => {
                        const isEven = index % 2 === 0;
                        return (
                          <tr 
                            key={`${item.id}-${index}`} 
                            className={`hover:bg-slate-900/40 transition-colors ${
                              isEven ? 'bg-slate-950/10' : 'bg-slate-950/30'
                            }`}
                          >
                            {/* ID */}
                            <td className="py-4 px-6 font-bold text-sky-400 whitespace-nowrap tracking-wide">
                              <div className="flex items-center gap-2">
                                <span>{item.idMaster}</span>
                                {(() => {
                                  const trend = getMasterTrend(item.idMaster);
                                  if (trend && trend.isWarning) {
                                    return (
                                      <span 
                                        className="bg-amber-500/15 text-amber-400 border border-amber-500/35 px-1.5 py-0.5 text-[8px] tracking-wide font-mono uppercase font-black flex items-center gap-1 animate-pulse"
                                        title={`Moyenne des 3 derniers calculs (${trend.averageMean}) a dérivé de ${trend.driftPercentage}% vers les limites. (Tolérances: LSL ${trend.lsl} / USL ${trend.usl})`}
                                      >
                                        <AlertTriangle className="h-2.5 w-2.5 text-amber-500 shrink-0" />
                                        <span>DÉRIVE {trend.driftPercentage}%</span>
                                      </span>
                                    );
                                  } else if (trend) {
                                    return (
                                      <span 
                                        className="bg-emerald-500/15 text-emerald-400 border border-emerald-500/25 px-1.5 py-0.5 text-[8px] tracking-wide font-mono uppercase font-black flex items-center gap-1"
                                        title={`Moyenne stable sur les 3 derniers calculs (dérive de ${trend.driftPercentage}%).`}
                                      >
                                        <CheckCircle className="h-2.5 w-2.5 text-emerald-500 shrink-0" />
                                        <span>STABLE</span>
                                      </span>
                                    );
                                  }
                                  return null;
                                })()}
                              </div>
                            </td>

                            {/* Tester */}
                            <td className="py-4 px-6 font-semibold text-slate-200">
                              {item.testeur}
                            </td>

                            {/* Product Info */}
                            <td className="py-4 px-6 max-w-[200px] truncate text-[11px]">
                              <div className="text-slate-300 font-bold">{item.refProduitMaster || <span className="text-slate-700">-</span>}</div>
                              {item.numSerieProduitMaster && (
                                <div className="text-[9px] text-slate-500 mt-0.5">SN: {item.numSerieProduitMaster}</div>
                              )}
                            </td>

                            {/* Card Info */}
                            <td className="py-4 px-6 max-w-[200px] truncate text-[11px]">
                              <div className="text-slate-300">{item.refCarteMaster || <span className="text-slate-700">-</span>}</div>
                              {item.numSerieCarteMaster && (
                                <div className="text-[9px] text-slate-500 mt-0.5">SN: {item.numSerieCarteMaster}</div>
                              )}
                            </td>

                            {/* Date creation */}
                            <td className="py-4 px-6 text-slate-400 whitespace-nowrap text-[11px]">
                              {item.dateCreation}
                            </td>

                            {/* Status badge */}
                            <td className="py-4 px-6 whitespace-nowrap">
                              {item.statut === "Active" ? (
                                <span className="bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 px-2 py-0.5 rounded-none font-bold text-[9px] uppercase tracking-wider inline-flex items-center gap-1">
                                  <span className="h-1.5 w-1.5 bg-emerald-500 inline-block shadow-[0_0_6px_#10b981]"></span>
                                  Active
                                </span>
                              ) : item.statut === "Obsolète" ? (
                                <span className="bg-amber-500/10 text-amber-400 border border-amber-500/20 px-2 py-0.5 rounded-none font-bold text-[9px] uppercase tracking-wider inline-flex items-center gap-1">
                                  <span className="h-1.5 w-1.5 bg-amber-500 inline-block"></span>
                                  Obsolète
                                </span>
                              ) : (
                                <span className="bg-rose-500/10 text-rose-400 border border-rose-500/20 px-2 py-0.5 rounded-none font-bold text-[9px] uppercase tracking-wider inline-flex items-center gap-1">
                                  <span className="h-1.5 w-1.5 bg-rose-500 inline-block"></span>
                                  Endommagée
                                </span>
                              )}
                            </td>

                            {/* Verification status badge */}
                            <td className="py-4 px-6 whitespace-nowrap text-[10px]">
                              {item.verif === "OK" ? (
                                <span className="text-emerald-400 bg-emerald-950/40 px-2 py-0.5 rounded-none border border-emerald-900 font-bold uppercase tracking-wider">OK</span>
                              ) : item.verif === "KO" ? (
                                <span className="text-rose-400 bg-rose-950/40 px-2 py-0.5 rounded-none border border-rose-900 font-bold uppercase tracking-wider">KO</span>
                              ) : (
                                <span className="text-slate-300 bg-slate-900/60 px-2 py-0.5 rounded-none border border-slate-800 font-bold uppercase tracking-wider">{item.verif}</span>
                              )}
                            </td>

                            {/* Actions */}
                            <td className="py-4 px-6 text-right whitespace-nowrap">
                              <div className="flex items-center justify-end gap-2">
                                {item.statut === "Active" && (
                                  <button
                                    onClick={() => handleSelectMasterForCalculation(item)}
                                    className="text-slate-950 bg-sky-500 hover:bg-sky-400 font-bold px-3 py-1 rounded-none border border-sky-400 flex items-center gap-1 text-[10px] uppercase tracking-wider transition-all"
                                    title="Calculer capabilité et répétabilité pour ce master"
                                  >
                                    <Calculator className="h-3 w-3" />
                                    <span>Calculer</span>
                                  </button>
                                )}

                                {true ? (
                                  <>
                                    <button
                                      onClick={() => handleEditMaster(item)}
                                      className="text-slate-400 hover:text-sky-400 p-1 rounded-none hover:bg-slate-900 border border-transparent hover:border-slate-800 transition-all cursor-pointer"
                                      title="Modifier la fiche Master (Modifier paramètres statut / verif)"
                                    >
                                      <Edit3 className="h-3.5 w-3.5" />
                                    </button>
                                    
                                    <button
                                      onClick={() => handleDeleteMaster(item.id)}
                                      className="text-slate-500 hover:text-rose-400 p-1 rounded-none hover:bg-slate-900 border border-transparent hover:border-slate-800 transition-all cursor-pointer"
                                      title="Supprimer la fiche Master"
                                    >
                                      <Trash2 className="h-3.5 w-3.5" />
                                    </button>
                                  </>
                                ) : (
                                  <div className="flex items-center text-slate-500 text-[9px] font-mono border border-slate-800 bg-slate-950/40 px-2.5 py-1 gap-1.5" title="Modification restreinte aux administrateurs (Lecture Seule)">
                                    <Lock className="h-3 w-3 text-amber-500/70" />
                                    <span>LECTURE SEULE</span>
                                  </div>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* General metadata verification logs - Geometric Balance Style */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 bg-slate-900/40 border border-slate-900 p-6 rounded-none shadow-md">
              <div className="space-y-2 border-r border-slate-900/40 pr-0 md:pr-6">
                <h4 className="font-black text-white text-xs uppercase tracking-wider font-display flex items-center gap-2">
                  <Info className="h-4 w-4 text-sky-500" />
                  Qu'est-ce qu'une pièce Master ?
                </h4>
                <p className="text-xs text-slate-400 leading-relaxed">
                  Dans les lignes de production électronique, une pièce <strong className="text-white">Master</strong> est un produit étalon dont les caractéristiques physiques sont connues avec une haute précision. Elle sert à vérifier périodiquement que les <strong className="text-white">bancs de tests (Testeurs)</strong> mesurent correctement et ne dérivent pas dans le temps (Répétabilité & Capabilité).
                </p>
              </div>

              <div className="space-y-2">
                <h4 className="font-black text-white text-xs uppercase tracking-wider font-display flex items-center gap-2">
                  <TrendingUp className="h-4 w-4 text-sky-500" />
                  Règle de Décision de Capabilité
                </h4>
                <p className="text-xs text-slate-400 leading-relaxed">
                  L'indice <strong className="text-sky-400 font-bold">Cpk</strong> mesure si la moyenne est centrée par rapport à l'intervalle de tolérance. Pour qu'un banc de test soit déclaré qualifié, il est d'usage industriel d'exiger un <strong className="text-emerald-400 font-bold">Cpk ≥ 1.33</strong> (ce qui correspond à un taux de rejet virtuel inférieur à 60 pièces par million).
                </p>
              </div>
            </div>

          </div>
        )}


        {/* ========================================================================= */}
        {/* TAB 2: CREER UN NOUVEAU MASTER (FORM) */}
        {activeTab === 'nouveau-master' && (
          <div className="max-w-3xl mx-auto animate-fadeIn">
            <div className="bg-slate-900/60 border border-slate-900 rounded-none overflow-hidden shadow-xl">
              
              <div className="bg-gradient-to-r from-emerald-950/40 via-slate-900/60 to-slate-900/20 border-b border-slate-900 p-6">
                <div className="flex items-center gap-4">
                  <div className="p-3 bg-emerald-500 text-slate-950 rounded-none shadow-[0_0_12px_rgba(16,185,129,0.2)]">
                    <PlusCircle className="h-5 w-5" />
                  </div>
                  <div>
                    <h3 className="font-black text-lg text-white font-display uppercase tracking-tight">Création d'une nouvelle pièce Master</h3>
                    <p className="text-xs text-slate-500 font-mono">Enregistrer une nouvelle fiche étalon pour l'un de vos bancs de test</p>
                  </div>
                </div>
              </div>

              <form onSubmit={handleCreateMaster} className="p-6 space-y-6">
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  
                  {/* Tester de destination */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Banc Test de Destination <span className="text-emerald-400">*</span>
                    </label>
                    <select
                      id="input_testeur"
                      value={newMaster.testeur}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, testeur: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-emerald-500 font-mono"
                    >
                      <option value="">-- Sélectionner un testeur existant --</option>
                      {uniqueTesters.map(tester => (
                        <option key={tester} value={tester}>{tester}</option>
                      ))}
                    </select>
                    
                    <div className="mt-2.5 flex items-center gap-2">
                      <span className="text-[9px] text-slate-500 font-mono uppercase tracking-wider whitespace-nowrap shrink-0">Ou saisir nouveau :</span>
                      <input
                        id="input_custom_tester"
                        type="text"
                        placeholder="Ex: UFT NEW_LINE"
                        onChange={(e) => setNewMaster(prev => ({ ...prev, testeur: e.target.value }))}
                        className="bg-slate-950 border border-slate-800 rounded-none px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-emerald-500 w-full font-mono"
                      />
                    </div>
                  </div>

                  {/* Identification Master */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Identification Master <span className="text-emerald-400">*</span>
                    </label>
                    <div className="flex gap-2">
                      <input
                        id="input_id_master"
                        type="text"
                        placeholder="Ex: MASTER_452"
                        value={newMaster.idMaster}
                        onChange={(e) => setNewMaster(prev => ({ ...prev, idMaster: e.target.value.toUpperCase() }))}
                        required
                        className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                      />
                      <button
                        type="button"
                        onClick={() => setNewMaster(prev => ({ ...prev, idMaster: getNextMasterId() }))}
                        className="bg-slate-950 hover:bg-slate-900 text-emerald-400 border border-slate-800 hover:border-slate-700 px-3 rounded-none text-[10px] font-mono uppercase font-bold tracking-wider transition-all"
                        title="Générer automatiquement le numéro séquentiel suivant"
                      >
                        Auto-ID
                      </button>
                    </div>
                    <span className="text-[9px] text-slate-500 block mt-1.5 font-mono">Format conseillé : <code>MASTER_XXX</code></span>
                  </div>

                  {/* Produit Reference */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Référence du Produit Master
                    </label>
                    <input
                      id="input_ref_produit"
                      type="text"
                      placeholder="Ex: SYT4501, AC964831C..."
                      value={newMaster.refProduitMaster}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, refProduitMaster: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Produit Serial Number */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      N° Série du Produit Master
                    </label>
                    <input
                      id="input_sn_produit"
                      type="text"
                      placeholder="Ex: 23089502530785"
                      value={newMaster.numSerieProduitMaster}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, numSerieProduitMaster: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Carte Reference */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Référence de la Carte Master
                    </label>
                    <input
                      id="input_ref_carte"
                      type="text"
                      placeholder="Ex: SY4501 A00"
                      value={newMaster.refCarteMaster}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, refCarteMaster: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Carte Serial Number */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      N° Série de la Carte Master
                    </label>
                    <input
                      id="input_sn_carte"
                      type="text"
                      placeholder="Ex: 21067564805283"
                      value={newMaster.numSerieCarteMaster}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, numSerieCarteMaster: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Status Selection */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Statut d'initialisation de la pièce
                    </label>
                    <select
                      id="input_status"
                      value={newMaster.statut}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, statut: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    >
                      <option value="Active">Active (Prête pour les bancs)</option>
                      <option value="Obsolète">Obsolète (Retirée de production)</option>
                      <option value="Endommagée">Endommagée (À réparer/Remplacer)</option>
                    </select>
                  </div>

                  {/* Metrology Verification */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Vérification Métrologique Initiale
                    </label>
                    <select
                      id="input_verif"
                      value={newMaster.verif}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, verif: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    >
                      <option value="OK">OK (Mesures validées)</option>
                      <option value="KO">KO (Dérive ou défaut physique)</option>
                      <option value="FIN DE VIE">FIN DE VIE</option>
                    </select>
                  </div>

                  {/* Date Picker */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Date d'enregistrement
                    </label>
                    <input
                      id="input_date"
                      type="text"
                      placeholder="Ex: 19/12/2024"
                      value={newMaster.dateCreation}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, dateCreation: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                  {/* Custom secondary comment */}
                  <div>
                    <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                      Commentaire / Note d'évolution (Optionnel)
                    </label>
                    <input
                      id="input_comment2"
                      type="text"
                      placeholder="Ex: Redesign ou modification d'indice..."
                      value={newMaster.commentaire2}
                      onChange={(e) => setNewMaster(prev => ({ ...prev, commentaire2: e.target.value }))}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-emerald-500"
                    />
                  </div>

                </div>

                {/* Main Comment */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">
                    Commentaires principaux / Spécifications du Testeur
                  </label>
                  <textarea
                    id="input_comment1"
                    rows={3}
                    placeholder="Ex: Spécificités géométriques du connecteur de test, limitations d'insertion..."
                    value={newMaster.commentaire1}
                    onChange={(e) => setNewMaster(prev => ({ ...prev, commentaire1: e.target.value }))}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none p-3 text-xs text-slate-100 placeholder-slate-600 focus:outline-none focus:border-emerald-500 font-mono"
                  />
                </div>

                {/* Form Buttons */}
                <div className="flex justify-end gap-3 border-t border-slate-900 pt-6">
                  <button
                    type="button"
                    onClick={() => setActiveTab('testeurs')}
                    className="bg-slate-950 hover:bg-slate-900 text-slate-300 border border-slate-800 rounded-none font-bold px-5 py-2.5 text-xs uppercase font-mono tracking-wider transition-colors"
                  >
                    Annuler
                  </button>
                  <button
                    id="btn_submit_master_form"
                    type="submit"
                    className="bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold px-6 py-2.5 rounded-none text-xs uppercase font-mono tracking-wider transition-all flex items-center gap-2 border border-emerald-400 shadow-[0_0_12px_rgba(16,185,129,0.15)]"
                  >
                    <CheckCircle className="h-4 w-4" /> Enregistrer le Master
                  </button>
                </div>

              </form>

            </div>
          </div>
        )}


        {/* ========================================================================= */}
        {/* TAB 3: CALCUL CAPABILITE & REPETABILITE                                    */}
        {/* ========================================================================= */}
        {activeTab === 'calcul' && (
          <div className="space-y-6 animate-fadeIn">
            <ActipaPanel calcMasterId={calcMasterId} masters={masters} session={session} />
          </div>
        )}
        {false && activeTab === 'calcul' && (
          <div className="space-y-6 animate-fadeIn">
            
            <div className="flex flex-col lg:flex-row gap-6">
              
              {/* Sidebar controls for inputs - Geometric Balance Style */}
              <div className="w-full lg:w-1/3 bg-slate-900/60 border border-slate-900 rounded-none p-5 space-y-6 shrink-0 shadow-lg font-mono">
                <div>
                  <div className="flex items-center justify-between mb-2">
                    <h4 className="font-bold text-white text-xs uppercase tracking-wider flex items-center gap-2 font-display">
                      <Settings className="h-4 w-4 text-sky-500" />
                      1. Paramètres Métrologiques
                    </h4>
                    <span className="text-[9px] text-slate-500 font-mono uppercase">Temps réel</span>
                  </div>
                  <p className="text-[11px] text-slate-400 leading-relaxed font-sans">Configurez les tolérances du banc pour calculer la capabilité machine.</p>
                </div>

                {/* Quick preset selector buttons */}
                <div className="space-y-2">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono">Essais prédéfinis (Presets) :</span>
                  <div className="grid grid-cols-2 gap-2">
                    <button
                      id="preset_5v"
                      onClick={() => loadPreset('5v')}
                      className="bg-slate-950 hover:bg-slate-900 border border-slate-800 hover:border-slate-700 px-2 py-2 rounded-none text-[9px] text-left text-slate-300 truncate uppercase font-bold tracking-wide"
                    >
                      🔌 Régul 5V (±0.15V)
                    </button>
                    <button
                      id="preset_conso"
                      onClick={() => loadPreset('conso')}
                      className="bg-slate-950 hover:bg-slate-900 border border-slate-800 hover:border-slate-700 px-2 py-2 rounded-none text-[9px] text-left text-slate-300 truncate uppercase font-bold tracking-wide"
                    >
                      🔋 Courant 24mA
                    </button>
                    <button
                      id="preset_rf"
                      onClick={() => loadPreset('rf')}
                      className="bg-slate-950 hover:bg-slate-900 border border-slate-800 hover:border-slate-700 px-2 py-2 rounded-none text-[9px] text-left text-slate-300 truncate uppercase font-bold tracking-wide"
                    >
                      📡 RF Smoove RTS
                    </button>
                    <button
                      id="preset_dim"
                      onClick={() => loadPreset('dim')}
                      className="bg-slate-950 hover:bg-slate-900 border border-slate-800 hover:border-slate-700 px-2 py-2 rounded-none text-[9px] text-left text-slate-300 truncate uppercase font-bold tracking-wide"
                    >
                      📐 Dimension 1.5mm
                    </button>
                  </div>
                </div>
                {/* ----------------- SELECTION DU MOTEUR DE CALCUL ----------------- */}
                <div className="pt-4 border-t border-slate-900 space-y-2">
                  <span className="text-[9px] font-bold text-slate-400 uppercase tracking-widest font-mono block">Moteur de Calcul :</span>
                  <div className="grid grid-cols-3 gap-1 p-0.5 bg-slate-950 border border-slate-900 rounded-none">
                    <button
                      type="button"
                      onClick={() => {
                        setCalcEngine('web');
                        setUseExternalCalc(false);
                      }}
                      className={`text-[8.5px] font-bold py-1.5 uppercase transition-all font-mono tracking-wider text-center cursor-pointer ${calcEngine === 'web' ? 'bg-sky-500 text-slate-950 font-black' : 'text-slate-500 hover:text-white hover:bg-slate-900/50'}`}
                    >
                      🌐 Web
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setCalcEngine('local');
                        setUseExternalCalc(true);
                      }}
                      className={`text-[8.5px] font-bold py-1.5 uppercase transition-all font-mono tracking-wider text-center cursor-pointer ${calcEngine === 'local' ? 'bg-sky-500 text-slate-950 font-black' : 'text-slate-500 hover:text-white hover:bg-slate-900/50'}`}
                    >
                      ⚙️ Local
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setCalcEngine('ai');
                        setUseExternalCalc(false);
                      }}
                      className={`text-[8.5px] font-bold py-1.5 uppercase transition-all font-mono tracking-wider text-center cursor-pointer ${calcEngine === 'ai' ? 'bg-indigo-500 text-slate-950 font-black' : 'text-slate-500 hover:text-white hover:bg-slate-900/50'}`}
                    >
                      🤖 IA Gemini
                    </button>
                  </div>
                </div>

                {/* Local Executable Configuration Panel */}
                {useExternalCalc && (
                  <div className="bg-slate-950/90 border border-sky-950/40 p-3.5 space-y-3.5 animate-fadeIn font-mono text-[10px] rounded-none">
                    <div className="flex items-center gap-1.5 text-sky-400 font-bold uppercase text-[9px] tracking-wider">
                      <Cpu className="h-3.5 w-3.5 text-sky-400 shrink-0" />
                      Paramètres Passerelle Locale
                    </div>

                    {/* Exe file path */}
                    <div>
                      <label className="block text-[8px] font-bold text-slate-500 uppercase tracking-widest mb-1">Chemin d'accès de l'exécutable (.exe)</label>
                      <input
                        type="text"
                        value={localExePath}
                        onChange={(e) => setLocalExePath(e.target.value)}
                        placeholder="C:\ACTIA\Metrologie\calcpk.exe"
                        className="w-full bg-slate-900 border border-slate-800 rounded-none px-2 py-1 font-mono text-[10px] text-slate-100 focus:outline-none focus:border-sky-500"
                      />
                    </div>

                    {/* Method Selector */}
                    <div>
                      <label className="block text-[8px] font-bold text-slate-500 uppercase tracking-widest mb-1">Liaison Browser ➔ Système</label>
                      <select
                        value={externalCalcMethod}
                        onChange={(e) => setExternalCalcMethod(e.target.value as 'api' | 'protocol')}
                        className="w-full bg-slate-900 border border-slate-800 rounded-none px-1.5 py-1 text-[10px] text-slate-300 font-mono focus:outline-none focus:border-sky-500"
                      >
                        <option value="api">Pont API local d'arrière-plan (Démon REST)</option>
                        <option value="protocol">Protocole URL direct (actiacalc://)</option>
                      </select>
                    </div>

                    {externalCalcMethod === 'api' ? (
                      <div className="space-y-2 pt-1 border-t border-slate-900">
                        <div>
                          <label className="block text-[8px] font-bold text-slate-500 uppercase tracking-widest mb-1">Adresse IP / Port du démon</label>
                          <input
                            type="text"
                            value={localBridgeUrl}
                            onChange={(e) => setLocalBridgeUrl(e.target.value)}
                            placeholder="http://localhost:5001/calculate"
                            className="w-full bg-slate-900 border border-slate-800 rounded-none px-2 py-1 font-mono text-[10px] text-sky-400 focus:outline-none focus:border-sky-500"
                          />
                        </div>

                        {/* Test connection row */}
                        <div className="flex items-center justify-between gap-2 pt-1">
                          <button
                            type="button"
                            onClick={handleTestLocalBridge}
                            disabled={isLocalTesting}
                            className="bg-slate-900 hover:bg-slate-800 border border-slate-800 disabled:opacity-50 text-slate-300 font-bold px-2 py-1 uppercase tracking-wider text-[8px] transition-all flex items-center gap-1 shrink-0"
                          >
                            <Link2 className={`h-3 w-3 ${isLocalTesting ? 'animate-spin text-sky-400' : ''}`} />
                            {isLocalTesting ? 'Test...' : 'Tester Liaison'}
                          </button>

                          <div className="flex items-center gap-1.5 text-[9px] font-bold uppercase select-none">
                            {localBridgeStatus === 'success' ? (
                              <>
                                <span className="h-2 w-2 rounded-full bg-emerald-500 animate-pulse"></span>
                                <span className="text-emerald-500">Actif (Connecté)</span>
                              </>
                            ) : localBridgeStatus === 'error' ? (
                              <>
                                <span className="h-2 w-2 rounded-full bg-rose-500"></span>
                                <span className="text-rose-500">Non détecté</span>
                              </>
                            ) : (
                              <>
                                <span className="h-2 w-2 rounded-full bg-slate-600"></span>
                                <span className="text-slate-500">Non testé</span>
                              </>
                            )}
                          </div>
                        </div>
                      </div>
                    ) : (
                      <p className="text-[9px] text-slate-500 leading-normal font-sans">
                        Le protocole personnalisé nécessite l'enregistrement préalable d'une clé de registre Windows pour associer <code className="text-sky-400 bg-slate-900 px-1 py-0.5 font-mono">actiacalc://</code> à votre exécutable.
                      </p>
                    )}
                  </div>
                )}

                {/* Form inputs */}
                <div className="space-y-3.5 pt-4 border-t border-slate-900">
                  
                  {/* Selected Master reference */}
                  <div>
                    <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Fiche Master affiliée</label>
                    <select
                      id="calc_select_master"
                      value={calcMasterId}
                      onChange={(e) => {
                        const m = masters.find(item => item.idMaster === e.target.value);
                        if (m) {
                          handleSelectMasterForCalculation(m);
                        } else {
                          setCalcMasterId(e.target.value);
                        }
                      }}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-sky-400 font-mono focus:outline-none focus:border-sky-500"
                    >
                      <option value="">-- Saisie manuelle sans master --</option>
                      {masters.map((m, index) => (
                        <option key={`${m.id}-${index}`} value={m.idMaster}>{m.idMaster} ({m.testeur})</option>
                      ))}
                    </select>
                  </div>

                  {/* Tester text */}
                  <div>
                    <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Nom du Testeur</label>
                    <select
                      id="calc_tester_select"
                      value={uniqueTesters.includes(calcTester) ? calcTester : ""}
                      onChange={(e) => {
                        if (e.target.value) {
                          handleCalcTesterChange(e.target.value);
                        }
                      }}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500 font-mono mb-2"
                    >
                      <option value="">-- Choisir un testeur existant --</option>
                      {uniqueTesters.map(t => (
                        <option key={t} value={t}>{t}</option>
                      ))}
                    </select>
                    
                    <input
                      id="calc_tester_input"
                      type="text"
                      value={calcTester}
                      onChange={(e) => handleCalcTesterChange(e.target.value)}
                      placeholder="Saisir ou modifier le nom du testeur"
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500"
                    />
                  </div>

                  {/* Characteristic */}
                  <div>
                    <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Caractéristique contrôlée</label>
                    <input
                      id="calc_characteristic_input"
                      type="text"
                      value={calcCharacteristic}
                      onChange={(e) => setCalcCharacteristic(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500"
                    />
                  </div>

                  {/* Sélecteur de Caractéristique Spéciale (SC) */}
                  <div>
                    <label className="block text-[9px] font-bold text-indigo-300 uppercase tracking-widest mb-1.5">Type de Caractéristique (SC)</label>
                    <select
                      id="select-sc"
                      value={calcIsSc}
                      onChange={(e) => setCalcIsSc(e.target.value as 'NON' | 'OUI')}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500"
                    >
                      <option value="NON">Standard (Seuil Cpk ≥ 1.33)</option>
                      <option value="OUI">Spéciale / Critique (Seuil Cpk ≥ 1.67)</option>
                    </select>
                  </div>

                  {/* Operator */}
                  <div>
                    <label className="block text-[9px] font-bold text-slate-400 uppercase tracking-widest mb-1.5">Technicien / Inspecteur</label>
                    <input
                      id="calc_operator_input"
                      type="text"
                      value={calcOperator}
                      onChange={(e) => setCalcOperator(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 focus:outline-none focus:border-sky-500"
                    />
                  </div>

                  {/* Specs limits */}
                  <div className="grid grid-cols-3 gap-2.5 pt-2">
                    <div>
                      <label className="block text-[8px] font-bold text-slate-400 uppercase tracking-widest mb-1">LSL (Min)</label>
                      <input
                        id="calc_lsl_input"
                        type="number"
                        step="any"
                        value={calcLsl}
                        onChange={(e) => setCalcLsl(parseFloat(e.target.value) || 0)}
                        className="w-full bg-slate-950 border border-slate-800 rounded-none px-2 py-1.5 text-xs font-mono text-rose-400 focus:outline-none focus:border-sky-500"
                      />
                    </div>
                    <div>
                      <label className="block text-[8px] font-bold text-slate-400 uppercase tracking-widest mb-1">Nominal</label>
                      <input
                        id="calc_nominal_input"
                        type="number"
                        step="any"
                        value={calcNominal}
                        onChange={(e) => setCalcNominal(parseFloat(e.target.value) || 0)}
                        className="w-full bg-slate-950 border border-slate-800 rounded-none px-2 py-1.5 text-xs font-mono text-sky-400 focus:outline-none focus:border-sky-500"
                      />
                    </div>
                    <div>
                      <label className="block text-[8px] font-bold text-slate-400 uppercase tracking-widest mb-1">USL (Max)</label>
                      <input
                        id="calc_usl_input"
                        type="number"
                        step="any"
                        value={calcUsl}
                        onChange={(e) => setCalcUsl(parseFloat(e.target.value) || 0)}
                        className="w-full bg-slate-950 border border-slate-800 rounded-none px-2 py-1.5 text-xs font-mono text-rose-400 focus:outline-none focus:border-sky-500"
                      />
                    </div>
                  </div>

                </div>

                {/* Measurements list textbox */}
                <div className="pt-4 border-t border-slate-900 space-y-4">
                  <div className="flex items-center justify-between">
                    <label className="block text-[10px] font-bold text-slate-300 uppercase tracking-wider font-display">
                      2. Relevés des Mesures
                    </label>
                    <button
                      type="button"
                      onClick={generateSyntheticData}
                      className="text-[9px] text-sky-500 bg-sky-950/30 hover:bg-sky-950/50 border border-sky-800/60 hover:border-sky-700 rounded-none px-3 py-1 font-mono uppercase tracking-wider font-bold transition-all flex items-center gap-1"
                    >
                      <Sparkles className="h-3 w-3" /> Générer 15 essais
                    </button>
                  </div>
                  
                  {/* Log File Drag & Drop Importer */}
                  <div className="space-y-2">
                    <div
                      onDragOver={(e) => {
                        e.preventDefault();
                        setIsDraggingLog(true);
                      }}
                      onDragLeave={() => setIsDraggingLog(false)}
                      onDrop={handleLogFileDrop}
                      onClick={() => document.getElementById("log_file_input")?.click()}
                      className={`border-2 border-dashed p-4 text-center cursor-pointer transition-all ${
                        isDraggingLog 
                          ? "border-sky-500 bg-sky-950/40 text-sky-400" 
                          : "border-slate-800 hover:border-slate-700 bg-slate-950/50 text-slate-400"
                      }`}
                    >
                      <input
                        id="log_file_input"
                        type="file"
                        accept=".log,.txt,.csv,.xml,.json"
                        multiple
                        className="hidden"
                        onChange={handleLogFileUpload}
                      />
                      <Upload className="h-5 w-5 mx-auto mb-2 text-sky-500 animate-pulse" />
                      <p className="text-[10px] font-bold uppercase tracking-wider text-slate-200">
                        Importer des fichiers log Master (SÉLECTION MULTIPLE)
                      </p>
                      <p className="text-[9px] text-slate-500 mt-1 font-sans">
                        Glissez-déposez vos fichiers de mesures (.log, .txt, .csv) ou cliquez pour parcourir (minimum 10 fichiers requis)
                      </p>
                    </div>

                    {/* Log Import details and actions */}
                    {logImportDetails && (
                      <div className="bg-slate-950 border border-slate-800 p-3 space-y-3 font-mono">
                        <div className="flex items-start justify-between border-b border-slate-900 pb-2">
                          <div className="space-y-1">
                            <span className="text-[9px] text-slate-500 block uppercase">
                              {logImportDetails.files && logImportDetails.files.length > 1 ? "Fichiers détectés" : "Fichier détecté"}
                            </span>
                            <span className="text-xs font-bold text-sky-400 block break-all">{logImportDetails.fileName}</span>
                          </div>
                          <button
                            type="button"
                            onClick={() => setLogImportDetails(null)}
                            className="text-slate-500 hover:text-rose-400 text-xs font-bold cursor-pointer"
                          >
                            ✕
                          </button>
                        </div>

                        {logImportDetails.files && logImportDetails.files.length > 1 && (
                          <div className="bg-slate-900/50 border border-slate-900/80 p-2 max-h-28 overflow-y-auto space-y-1.5 text-[10px] rounded-none">
                            <span className="text-[8px] font-bold text-slate-500 uppercase tracking-wider block mb-1">Détails par fichier :</span>
                            {logImportDetails.files.map((f, idx) => (
                              <div key={idx} className="flex justify-between items-center text-slate-300 border-b border-slate-950 pb-1 last:border-0 last:pb-0">
                                <span className="truncate max-w-[170px] text-slate-400 font-sans">📄 {f.fileName}</span>
                                <span className="text-sky-400 font-bold shrink-0">{f.valuesFound.length} val.</span>
                              </div>
                            ))}
                          </div>
                        )}

                        <div className="grid grid-cols-2 gap-2 text-[10px] text-slate-400">
                          <div>Fichiers lus: <strong className="text-white">{logImportDetails.files ? logImportDetails.files.length : 1}</strong></div>
                          <div>Mesures cumulées: <strong className="text-emerald-400">{logImportDetails.valuesFound.length}</strong></div>
                          {logImportDetails.characteristic && (
                            <div className="col-span-2 truncate">Caractéristique: <strong className="text-sky-400">{logImportDetails.characteristic}</strong></div>
                          )}
                          {(logImportDetails.nominal !== undefined || logImportDetails.lsl !== undefined || logImportDetails.usl !== undefined) && (
                            <div className="col-span-2 text-[9px] text-slate-500">
                              Paramètres détectés : 
                              {logImportDetails.nominal !== undefined && ` Nominal: ${logImportDetails.nominal}`}
                              {logImportDetails.lsl !== undefined && ` Min: ${logImportDetails.lsl}`}
                              {logImportDetails.usl !== undefined && ` Max: ${logImportDetails.usl}`}
                            </div>
                          )}
                        </div>

                        {/* Preview of first few values */}
                        <div className="bg-slate-900/60 p-2 text-[9px] text-slate-400 max-h-16 overflow-y-auto break-words select-none leading-relaxed rounded-none">
                          Aperçu des valeurs extraites : <span className="text-white font-bold">{logImportDetails.valuesFound.slice(0, 15).join(", ")}{logImportDetails.valuesFound.length > 15 ? "..." : ""}</span>
                        </div>

                        <div className="flex gap-2 text-[10px]">
                          <button
                            type="button"
                            onClick={() => handleInjectLogValues('replace')}
                            className="flex-1 bg-sky-500 hover:bg-sky-400 text-slate-950 font-bold py-2 uppercase text-center transition-colors cursor-pointer"
                          >
                            Remplacer les mesures
                          </button>
                          <button
                            type="button"
                            onClick={() => handleInjectLogValues('append')}
                            className="flex-1 bg-slate-900 hover:bg-slate-800 border border-slate-800 text-slate-300 font-bold py-2 uppercase text-center transition-colors cursor-pointer"
                          >
                            Ajouter aux mesures
                          </button>
                        </div>
                      </div>
                    )}
                  </div>

                  <p className="text-[10px] text-slate-500 leading-normal font-sans">
                    Saisie manuelle ou collage de données (colonne Excel, valeurs séparées par espace/retour chariot) :
                  </p>
                  
                  <textarea
                    id="calc_measurements_input"
                    rows={6}
                    value={measurementsInput}
                    onChange={(e) => setMeasurementsInput(e.target.value)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none p-2.5 font-mono text-xs text-slate-200 focus:outline-none focus:border-sky-500"
                    placeholder="Entrez ou modifiez vos mesures ici..."
                  />

                  <div className="flex justify-between items-center text-[10px] text-slate-500 font-mono pt-1">
                    <span>Nombre de relevés analysés :</span>
                    <span className="text-sky-400 font-bold bg-slate-950 px-2 py-0.5 rounded-none border border-slate-900">
                      {measurementsArray.length} valeurs
                    </span>
                  </div>

                  {/* External Background calculation primary trigger */}
                  {calcEngine === 'local' && (
                    <div className="pt-4 border-t border-slate-900 space-y-2.5 animate-fadeIn">
                      {externalCalcMethod === 'api' ? (
                        <button
                          type="button"
                          onClick={handleExecuteExternalCalc}
                          disabled={isLocalCalculating || measurementsArray.length === 0}
                          className="w-full bg-sky-500 hover:bg-sky-400 disabled:bg-slate-800 disabled:text-slate-500 text-slate-950 font-black py-3 rounded-none uppercase font-mono tracking-wider text-xs transition-all flex items-center justify-center gap-2 border border-sky-400 shadow-[0_0_15px_rgba(14,165,233,0.2)] cursor-pointer"
                        >
                          <Cpu className={`h-4 w-4 ${isLocalCalculating ? 'animate-spin' : ''}`} />
                          {isLocalCalculating ? "Calcul Local en cours..." : "Lancer le Calcul Externe (.exe)"}
                        </button>
                      ) : (
                        <button
                          type="button"
                          onClick={handleTriggerCustomProtocol}
                          disabled={measurementsArray.length === 0}
                          className="w-full bg-emerald-500 hover:bg-emerald-400 disabled:opacity-50 text-slate-950 font-black py-3 rounded-none uppercase font-mono tracking-wider text-xs transition-all flex items-center justify-center gap-2 border border-emerald-400 shadow-[0_0_15px_rgba(16,185,129,0.2)] cursor-pointer"
                        >
                          <ExternalLink className="h-4 w-4" />
                          Ouvrir via actiacalc://
                        </button>
                      )}
                      <p className="text-[9px] text-slate-500 leading-normal font-sans text-center">
                        Exécute <code className="text-slate-400 font-mono bg-slate-950 px-1 py-0.5">{localExePath.split('\\').pop() || 'calcpk.exe'}</code> avec les {measurementsArray.length} mesures en tâche d'arrière-plan.
                      </p>
                    </div>
                  )}

                  {/* Gemini AI calculation primary trigger */}
                  {calcEngine === 'ai' && (
                    <div className="pt-4 border-t border-slate-900 space-y-2.5 animate-fadeIn">
                      <button
                        type="button"
                        onClick={handleExecuteAiCalc}
                        disabled={isAiCalculating || measurementsArray.length === 0}
                        className="w-full bg-indigo-600 hover:bg-indigo-500 disabled:bg-slate-800 disabled:text-slate-500 text-white font-black py-3 rounded-none uppercase font-mono tracking-wider text-xs transition-all flex items-center justify-center gap-2 border border-indigo-400 shadow-[0_0_15px_rgba(99,102,241,0.2)] cursor-pointer"
                      >
                        <Sparkles className={`h-4 w-4 ${isAiCalculating ? 'animate-spin' : ''}`} />
                        {isAiCalculating ? "Calcul IA en cours..." : "Lancer le Calcul Métrologique IA"}
                      </button>
                      <p className="text-[9px] text-slate-500 leading-normal font-sans text-center font-mono">
                        Exécute une analyse statistique avancée avec l'intelligence artificielle Gemini. Détecte et nettoie automatiquement les valeurs aberrantes (outliers) pour une capabilité de précision optimale.
                      </p>
                    </div>
                  )}

                </div>

              </div>

              {/* Main capability results screen - Geometric Balance Style */}
              <div className="flex-1 bg-slate-900/40 border border-slate-900 rounded-none p-6 space-y-6 shadow-lg">
                
                {/* Section Header */}
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 border-b border-slate-900 pb-4">
                  <div>
                    <h3 className="font-bold text-white text-sm uppercase tracking-wider flex items-center gap-2 font-display">
                      <TrendingUp className="h-5 w-5 text-sky-500" />
                      Rapport Qualité de Répétabilité & Capabilité
                    </h3>
                    <p className="text-[11px] text-slate-400 font-sans leading-relaxed">Calculs d'indices machine Cp, Cpk et d'étendue d'essais (Normes ISO/TS)</p>
                  </div>

                  <div className="flex items-center gap-2">
                    {calcMasterId && (
                      <button
                        id="btn_save_calc_history"
                        onClick={handleSaveCalculation}
                        className="text-xs bg-sky-500 hover:bg-sky-400 text-slate-950 px-4 py-2 rounded-none font-bold font-mono uppercase tracking-wider flex items-center gap-1.5 transition-all cursor-pointer"
                        title="Enregistrer ce calcul de capabilité dans l'historique du master pour l'analyse des tendances SPC"
                      >
                        <TrendingUp className="h-4 w-4" />
                        Enregistrer le Calcul
                      </button>
                    )}
                    <button
                      id="btn_print_report"
                      onClick={handlePrint}
                      className="text-xs bg-slate-950 hover:bg-slate-900 border border-slate-800 text-slate-200 px-4 py-2 rounded-none font-bold font-mono uppercase tracking-wider flex items-center gap-1.5 transition-all cursor-pointer"
                    >
                      <FileText className="h-4 w-4 text-sky-500" />
                      Imprimer
                    </button>
                    <button
                      id="btn_export_pdf"
                      onClick={handleExportPDF}
                      disabled={isExportingPDF}
                      className="text-xs bg-emerald-600 hover:bg-emerald-500 disabled:opacity-50 text-white px-4 py-2 rounded-none font-bold font-mono uppercase tracking-wider flex items-center gap-1.5 transition-all cursor-pointer"
                      title="Générer et télécharger le rapport officiel complet de capabilité au format PDF"
                    >
                      {isExportingPDF ? (
                        <RefreshCw className="h-4 w-4 animate-spin text-white" />
                      ) : (
                        <Download className="h-4 w-4 text-white" />
                      )}
                      <span>{isExportingPDF ? "Export..." : "Exporter PDF"}</span>
                    </button>
                  </div>
                </div>

                {/* Main statistics dashboard results */}
                {!activeStats ? (
                  <div className="p-16 text-center space-y-3 font-mono">
                    <Calculator className="h-12 w-12 text-slate-700 mx-auto animate-pulse" />
                    <p className="text-slate-300 font-bold uppercase text-xs tracking-wider">En attente de données d'essais</p>
                    <p className="text-[11px] text-slate-500 max-w-sm mx-auto font-sans leading-normal">
                      Veuillez saisir des relevés numériques dans la zone d'essais à gauche, importer un fichier log, ou utiliser l'exécution d'application externe locale.
                    </p>
                  </div>
                ) : (
                  <div className="space-y-6 font-mono">
                    
                    {/* Gauge Index Display and Decision Banner */}
                    <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
                      
                      {/* Classification Badge Card */}
                      <div className="lg:col-span-2 rounded-none p-5 border flex flex-col justify-between bg-slate-950/40 border-slate-900">
                        <div className="space-y-2">
                          <span className="text-[9px] font-bold tracking-widest text-slate-500 uppercase">Évaluation du Banc de Test</span>
                          <h4 className="text-sm font-bold text-white flex items-center gap-2 font-display">
                            {activeStats.cpk >= (calcIsSc === 'OUI' ? 1.67 : 1.33) ? (
                              <>
                                <CheckCircle className="h-5 w-5 text-emerald-500 shrink-0" />
                                <span className="text-emerald-500 uppercase">CAPABLE & VALIDÉ (Excellent)</span>
                              </>
                            ) : activeStats.cpk >= (calcIsSc === 'OUI' ? 1.33 : 1.0) ? (
                              <>
                                <AlertTriangle className="h-5 w-5 text-amber-500 shrink-0" />
                                <span className="text-amber-500 uppercase">MARGINAL (À surveiller)</span>
                              </>
                            ) : (
                              <>
                                <XCircle className="h-5 w-5 text-rose-500 shrink-0" />
                                <span className="text-rose-500 font-bold uppercase">INSUFFISANT / REJETÉ (Non capable)</span>
                              </>
                            )}
                          </h4>
                          <p className="text-xs text-slate-400 leading-relaxed font-sans">
                            {activeStats.diagnostic || (activeStats.cpk >= (calcIsSc === 'OUI' ? 1.67 : 1.33) 
                              ? `Le processus possède une marge de sécurité idéale. Le banc de test ${calcTester} est parfaitement apte à discriminer les pièces conformes.`
                              : activeStats.cpk >= (calcIsSc === 'OUI' ? 1.33 : 1.0)
                              ? `Le banc est capable mais frôle la zone de dérive. Des actions de maintenance métrologique préventive ou un ré-étalonnage sont préconisés.`
                              : `Le banc présente une dispersion trop élevée ou sa moyenne est décentrée. La capabilité actuelle (${activeStats.cpk.toFixed(2)}) est inférieure au seuil critique de ${(calcIsSc === 'OUI' ? 1.67 : 1.33).toFixed(2)}.`
                            )}
                          </p>
                        </div>
                        
                        <div className="pt-3.5 mt-3 border-t border-slate-900 flex items-center justify-between text-[10px] font-mono text-slate-400">
                          <span>Seuil qualitatif exigé : <strong className="text-sky-500">Cpk ≥ {calcIsSc === 'OUI' ? '1.67' : '1.33'}</strong></span>
                          <span>Cpk Réel : <strong className={activeStats.cpk >= (calcIsSc === 'OUI' ? 1.67 : 1.33) ? 'text-emerald-500' : activeStats.cpk >= (calcIsSc === 'OUI' ? 1.33 : 1.0) ? 'text-amber-500' : 'text-rose-500'}>{activeStats.cpk.toFixed(3)}</strong></span>
                        </div>
                      </div>

                      {/* Speedometer CPK Meter Dial */}
                      <div className="bg-slate-950/40 border border-slate-900 rounded-none p-4 flex flex-col items-center justify-center text-center">
                        <span className="text-[9px] font-bold tracking-widest text-slate-500 uppercase mb-2">Proportion Cpk</span>
                        <div className="relative h-20 w-36 flex items-end justify-center overflow-hidden">
                          {/* Half circle track */}
                          <svg className="absolute w-36 h-20" viewBox="0 0 100 50">
                            {/* Gray background arc */}
                            <path d="M 10 50 A 40 40 0 0 1 90 50" fill="none" stroke="#0f172a" strokeWidth="12" strokeLinecap="none" />
                            {/* colored active arc based on cpk (cap at 2.0 max for display) */}
                            <path 
                              d="M 10 50 A 40 40 0 0 1 90 50" 
                              fill="none" 
                              stroke={activeStats.cpk >= (calcIsSc === 'OUI' ? 1.67 : 1.33) ? '#10b981' : activeStats.cpk >= (calcIsSc === 'OUI' ? 1.33 : 1.0) ? '#f59e0b' : '#f43f5e'} 
                              strokeWidth="12" 
                              strokeLinecap="none"
                              strokeDasharray="125.6"
                              strokeDashoffset={125.6 - (Math.min(activeStats.cpk, 2.0) / 2.0) * 125.6}
                              className="transition-all duration-1000 ease-out"
                            />
                          </svg>
                          <div className="relative text-center z-10">
                            <span className="text-2xl font-black font-mono tracking-tight text-white">{activeStats.cpk.toFixed(2)}</span>
                            <div className="text-[9px] text-slate-500 uppercase font-mono font-bold">Index Cpk</div>
                          </div>
                        </div>
                        <div className="text-[9px] text-slate-500 mt-2 font-mono">
                          {activeStats.cpk >= 2.0 ? "Précision maximale" : `Cp calculé : ${activeStats.cp.toFixed(2)}`}
                        </div>
                      </div>

                      {/* D3 Radar Chart */}
                      <div className="bg-slate-950/40 border border-slate-900 rounded-none p-4 flex flex-col items-center justify-center text-center">
                        <span className="text-[9px] font-bold tracking-widest text-slate-500 uppercase mb-2">Points Critiques (Radar D3)</span>
                        <div className="w-full flex items-center justify-center">
                          <RadarChart
                            cp={activeStats.cp}
                            cpk={activeStats.cpk}
                            centering={(2 * (activeStats.mean - calcNominal)) / (calcUsl - calcLsl)}
                            sigma={activeStats.stdDev}
                            range={activeStats.range}
                            lsl={calcLsl}
                            usl={calcUsl}
                            nominal={calcNominal}
                            width={160}
                            height={150}
                            theme="dark"
                          />
                        </div>
                      </div>

                    </div>

                    {/* Quality KPI Cards - Geometric Balance style with left accent borders */}
                    <div className="grid grid-cols-2 md:grid-cols-3 xl:grid-cols-6 gap-3 pt-2">
                      
                      {/* Moyenne */}
                      <div className="bg-slate-950/40 p-3.5 rounded-none border-l-2 border-l-sky-500 border-y border-r border-slate-900 text-center">
                        <div className="text-[9px] font-mono text-slate-500 uppercase tracking-wider">Moyenne (X-bar)</div>
                        <div className="text-base font-bold font-mono text-white mt-1">{activeStats.mean.toFixed(4)}</div>
                        <div className="text-[8px] text-slate-600 font-mono mt-0.5">Cible: {calcNominal}</div>
                      </div>

                      {/* Ecart type */}
                      <div className="bg-slate-950/40 p-3.5 rounded-none border-l-2 border-l-sky-400 border-y border-r border-slate-900 text-center">
                        <div className="text-[9px] font-mono text-slate-500 uppercase tracking-wider">Écart-Type (σ)</div>
                        <div className="text-base font-bold font-mono text-sky-400 mt-1">{activeStats.stdDev.toFixed(5)}</div>
                        <div className="text-[8px] text-slate-600 font-mono mt-0.5">Dispersion machine</div>
                      </div>

                      {/* Etendue */}
                      <div className="bg-slate-950/40 p-3.5 rounded-none border-l-2 border-l-slate-400 border-y border-r border-slate-900 text-center">
                        <div className="text-[9px] font-mono text-slate-500 uppercase tracking-wider">Étendue (R)</div>
                        <div className="text-base font-bold font-mono text-slate-300 mt-1">{activeStats.range.toFixed(4)}</div>
                        <div className="text-[8px] text-slate-600 font-mono mt-0.5">Max - Min</div>
                      </div>

                      {/* Répétabilité */}
                      <div className="bg-slate-950/40 p-3.5 rounded-none border-l-2 border-l-rose-500 border-y border-r border-slate-900 text-center">
                        <div className="text-[9px] font-mono text-slate-500 uppercase tracking-wider">Répétabilité (6σ)</div>
                        <div className="text-base font-bold font-mono text-rose-400 mt-1">{activeStats.repeatability.toFixed(4)}</div>
                        <div className="text-[8px] text-slate-600 font-mono mt-0.5">Dispersion 99.73%</div>
                      </div>

                      {/* Cp */}
                      <div className="bg-slate-950/40 p-3.5 rounded-none border-l-2 border-l-slate-300 border-y border-r border-slate-900 text-center">
                        <div className="text-[9px] font-mono text-slate-500 uppercase tracking-wider">Potentiel Cp</div>
                        <div className="text-base font-bold font-mono text-white mt-1">{activeStats.cp.toFixed(3)}</div>
                        <div className="text-[8px] text-slate-600 font-mono mt-0.5">Aptitude potentielle</div>
                      </div>

                      {/* Cpk */}
                      <div className="bg-slate-950/40 p-3.5 rounded-none border-l-2 border-l-emerald-500 border-y border-r border-slate-900 text-center">
                        <div className="text-[9px] font-mono text-slate-500 uppercase tracking-wider">Ajusté Cpk</div>
                        <div className="text-base font-bold font-mono text-emerald-400 mt-1">{activeStats.cpk.toFixed(3)}</div>
                        <div className="text-[8px] text-slate-600 font-mono mt-0.5">Limité par {activeStats.cpkType}</div>
                      </div>

                    </div>

                    {/* SPC Trend Analysis card */}
                    {calcMasterId && (() => {
                      const trend = getMasterTrend(calcMasterId);
                      if (!trend) {
                        return (
                          <div className="bg-slate-950/20 border border-slate-900 p-4 rounded-none font-mono text-slate-500 text-[10px] uppercase flex items-center gap-2">
                            <TrendingUp className="h-4 w-4 text-slate-700" />
                            <span>Historique SPC insuffisant (&lt; 3 calculs de capabilité enregistrés pour {calcMasterId}) pour analyser les tendances. Cliquez sur "Enregistrer le Calcul" pour enrichir la base de mesures.</span>
                          </div>
                        );
                      }

                      return (
                        <div className={`p-5 border rounded-none ${trend.isWarning ? 'bg-amber-950/20 border-amber-500/40' : 'bg-slate-950/40 border-slate-900'} font-mono space-y-3`}>
                          <div className="flex items-center justify-between">
                            <span className="text-[10px] font-bold tracking-widest text-slate-400 uppercase flex items-center gap-1.5">
                              <TrendingUp className="h-4 w-4 text-sky-500" />
                              Analyse de Tendance & Dérive SPC
                            </span>
                            {trend.isWarning ? (
                              <span className="bg-amber-500/10 text-amber-400 border border-amber-500/30 px-2 py-0.5 text-[9px] uppercase font-black flex items-center gap-1 animate-pulse">
                                <AlertTriangle className="h-3.5 w-3.5 text-amber-500 shrink-0" />
                                <span>⚠️ DÉRIVE DÉTECTÉE &gt; 5%</span>
                              </span>
                            ) : (
                              <span className="bg-emerald-500/10 text-emerald-400 border border-emerald-500/20 px-2 py-0.5 text-[9px] uppercase font-black flex items-center gap-1">
                                <CheckCircle className="h-3.5 w-3.5 text-emerald-500 shrink-0" />
                                <span>STABLE (Sous Contrôle)</span>
                              </span>
                            )}
                          </div>

                          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                            <div>
                              <div className="text-[9px] text-slate-500 uppercase">Moyenne des 3 derniers calculs</div>
                              <div className={`text-base font-bold ${trend.isWarning ? 'text-amber-400' : 'text-slate-200'}`}>
                                {trend.averageMean}
                              </div>
                              <div className="text-[9px] text-slate-600 mt-0.5">Valeur nominale : {trend.nominal}</div>
                            </div>

                            <div>
                              <div className="text-[9px] text-slate-500 uppercase">Pourcentage de dérive</div>
                              <div className={`text-base font-bold ${trend.isWarning ? 'text-amber-400' : 'text-emerald-400'}`}>
                                {trend.driftPercentage}%
                              </div>
                              <div className="text-[9px] text-slate-600 mt-0.5">Seuil critique : &gt; 5% de la bande</div>
                            </div>

                            <div>
                              <div className="text-[9px] text-slate-500 uppercase">Historique des 3 moyennes</div>
                              <div className="text-slate-300 font-bold tracking-wide flex items-center gap-1 mt-1 text-[11px]">
                                {trend.means.map((m, idx) => (
                                  <React.Fragment key={idx}>
                                    {idx > 0 && <span className="text-slate-600">➔</span>}
                                    <span className={trend.isWarning && idx === 2 ? 'text-amber-400 font-black' : 'text-slate-400'}>{m}</span>
                                  </React.Fragment>
                                ))}
                              </div>
                              <div className="text-[9px] text-slate-600 mt-0.5 font-sans">Bande (LSL-USL) : {trend.tolerance.toFixed(3)}</div>
                            </div>
                          </div>

                          <p className="text-[11px] text-slate-400 font-sans leading-relaxed pt-1.5 border-t border-slate-900/60">
                            {trend.isWarning 
                              ? `Avertissement métrologique : La dérive moyenne du master ${calcMasterId} dépasse 5% de l'intervalle de tolérance (${(trend.tolerance).toFixed(3)}). Le banc de test présente un biais systématique de mesure. Une inspection matérielle ou un étalonnage complémentaire est fortement recommandé.`
                              : `Le comportement du master ${calcMasterId} est stable. La moyenne cumulée des mesures reste parfaitement centrée sur la valeur cible dans les limites admissibles.`
                            }
                          </p>
                        </div>
                      );
                    })()}

                    {/* Premium Custom Interactive SVG Chart */}
                    <div className="bg-slate-950/40 border border-slate-900 rounded-none p-5 space-y-4">
                      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2">
                        <div>
                          <h4 className="font-bold text-white text-xs uppercase tracking-wider font-mono">
                            Graphique d'évolution des mesures par rapport aux limites
                          </h4>
                          <p className="text-[10px] text-slate-400 font-sans leading-relaxed">Aperçu chronologique de la dispersion des {activeStats.n} essais</p>
                        </div>

                        {/* Legend */}
                        <div className="flex flex-wrap gap-x-4 gap-y-1 text-[9px] font-mono uppercase">
                          <span className="flex items-center gap-1"><span className="h-0.5 w-3 bg-rose-500 border border-rose-500 border-dashed inline-block"></span> USL / LSL</span>
                          <span className="flex items-center gap-1"><span className="h-0.5 w-3 bg-sky-500 border border-sky-500 border-dashed inline-block"></span> Nominal ({calcNominal})</span>
                          <span className="flex items-center gap-1"><span className="h-0.5 w-3 bg-emerald-500 inline-block"></span> Moyenne ({activeStats.mean.toFixed(3)})</span>
                        </div>
                      </div>

                      {/* Render custom high fidelity chart */}
                      <div className="relative w-full h-64 bg-slate-950 border border-slate-900 rounded-none p-2 overflow-hidden">
                        
                        {/* Custom SVG Drawing */}
                        <svg className="w-full h-full" viewBox="0 0 600 220" preserveAspectRatio="none">
                          {/* Grid Lines */}
                          <line x1="40" y1="20" x2="580" y2="20" stroke="#1e293b" strokeDasharray="2" />
                          <line x1="40" y1="70" x2="580" y2="70" stroke="#1e293b" strokeDasharray="2" />
                          <line x1="40" y1="120" x2="580" y2="120" stroke="#1e293b" strokeDasharray="2" />
                          <line x1="40" y1="170" x2="580" y2="170" stroke="#1e293b" strokeDasharray="2" />

                          {/* Compute scale helpers */}
                          {(() => {
                            const arr = measurementsArray;
                            const n = arr.length;
                            
                            // Determine min/max boundary of chart based on specs and measurements
                            const maxLimit = Math.max(calcUsl, ...arr);
                            const minLimit = Math.min(calcLsl, ...arr);
                            const margin = (maxLimit - minLimit) * 0.15 || 1.0;
                            const topY = maxLimit + margin;
                            const bottomY = minLimit - margin;
                            const rangeY = topY - bottomY;

                            // Pixel converters
                            const getX = (index: number) => {
                              if (n <= 1) return 310;
                              return 40 + (index / (n - 1)) * (580 - 40);
                            };
                            const getY = (val: number) => {
                              // map topY -> 20px, bottomY -> 200px
                              return 20 + ((topY - val) / rangeY) * (200 - 20);
                            };

                            // Points coordinates
                            const points = arr.map((val, idx) => ({
                              x: getX(idx),
                              y: getY(val),
                              val,
                              idx
                            }));

                            // Draw USL and LSL zone fills
                            const uslY = getY(calcUsl);
                            const lslY = getY(calcLsl);
                            const nomY = getY(calcNominal);
                            const meanY = getY(activeStats.mean);

                            return (
                              <>
                                {/* Safe tolerance zone fill (soft teal/green) */}
                                <rect 
                                  x="40" 
                                  y={Math.min(uslY, lslY)} 
                                  width="540" 
                                  height={Math.abs(lslY - uslY)} 
                                  fill="#10b981" 
                                  fillOpacity="0.03" 
                                />

                                {/* Out of specs zone fills */}
                                <rect x="40" y="0" width="540" height={Math.max(0, uslY)} fill="#f43f5e" fillOpacity="0.05" />
                                <rect x="40" y={lslY} width="540" height={Math.max(0, 220 - lslY)} fill="#f43f5e" fillOpacity="0.05" />

                                {/* Reference line USL */}
                                <line x1="40" y1={uslY} x2="580" y2={uslY} stroke="#f43f5e" strokeWidth="1.5" strokeDasharray="4 4" />
                                <text x="45" y={uslY - 5} fill="#f43f5e" className="text-[8px] font-mono font-bold">USL: {calcUsl}</text>

                                {/* Reference line LSL */}
                                <line x1="40" y1={lslY} x2="580" y2={lslY} stroke="#f43f5e" strokeWidth="1.5" strokeDasharray="4 4" />
                                <text x="45" y={lslY + 11} fill="#f43f5e" className="text-[8px] font-mono font-bold">LSL: {calcLsl}</text>

                                {/* Nominal line */}
                                <line x1="40" y1={nomY} x2="580" y2={nomY} stroke="#3b82f6" strokeWidth="1" strokeDasharray="3 3" />
                                <text x="540" y={nomY - 5} fill="#3b82f6" className="text-[8px] font-mono">Nom: {calcNominal}</text>

                                {/* Process Mean line */}
                                <line x1="40" y1={meanY} x2="580" y2={meanY} stroke="#10b981" strokeWidth="1.5" />
                                <text x="520" y={meanY + 11} fill="#10b981" className="text-[8px] font-mono font-bold">Moyenne: {activeStats.mean.toFixed(3)}</text>

                                {/* Plot line connecting points */}
                                {points.length > 1 && (
                                  <path
                                    d={points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ')}
                                    fill="none"
                                    stroke="#14b8a6"
                                    strokeWidth="2"
                                    className="opacity-80"
                                  />
                                )}

                                {/* Draw individual point nodes */}
                                {points.map((p) => {
                                  const isOut = p.val > calcUsl || p.val < calcLsl;
                                  return (
                                    <g key={p.idx}>
                                      <circle
                                        cx={p.x}
                                        cy={p.y}
                                        r={isOut ? "5" : "4"}
                                        fill={isOut ? "#f43f5e" : "#14b8a6"}
                                        stroke="#0f172a"
                                        strokeWidth="1.5"
                                        className="cursor-pointer hover:r-6 hover:stroke-teal-300 transition-all"
                                      />
                                      <text 
                                        x={p.x} 
                                        y={p.y - 8} 
                                        textAnchor="middle" 
                                        fill="#94a3b8" 
                                        className="text-[7px] font-mono"
                                      >
                                        #{p.idx + 1}
                                      </text>
                                    </g>
                                  );
                                })}
                              </>
                            );
                          })()}
                        </svg>

                        {/* Chart background indicator */}
                        <div className="absolute top-2 left-2 bg-slate-950/90 text-slate-500 text-[8px] font-mono px-2 py-0.5 rounded-none border border-slate-900">
                          Abscisse : Numéro d'essai d'échantillonnage
                        </div>
                      </div>
                    </div>

                    {/* Measurements summary table */}
                    <div className="bg-slate-950/40 border border-slate-900 rounded-none p-4">
                      <h4 className="font-bold text-slate-300 text-xs uppercase tracking-wider font-display mb-3 flex items-center gap-2">
                        <FileText className="h-4 w-4 text-sky-500" />
                        Détail des essais de mesure (Mesures individuelles)
                      </h4>
                      <div className="grid grid-cols-3 sm:grid-cols-5 gap-3 max-h-48 overflow-y-auto pr-2">
                        {measurementsArray.map((v, i) => {
                          const deviation = v - calcNominal;
                          const isOut = v > calcUsl || v < calcLsl;
                          return (
                            <div 
                              key={i} 
                              className={`p-2 rounded-none border text-center font-mono text-xs ${
                                isOut 
                                  ? 'bg-rose-950/60 border-rose-900 text-rose-400' 
                                  : 'bg-slate-950 border-slate-900 text-slate-300'
                              }`}
                            >
                              <div className="text-[9px] text-slate-500 uppercase font-bold">Essai #{i+1}</div>
                              <div className="text-sm font-black my-0.5">{v.toFixed(3)}</div>
                              <div className={`text-[8px] ${deviation >= 0 ? 'text-emerald-500' : 'text-rose-400'}`}>
                                {deviation >= 0 ? `+${deviation.toFixed(3)}` : deviation.toFixed(3)}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>

                    {/* Audit Log / Certificate signoff template */}
                    <div className="border-t border-slate-900 pt-6 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 font-mono text-[9px] text-slate-500 uppercase tracking-wider">
                      <div>ID Rapport : <span className="text-slate-400 font-bold">CAP-{calcMasterId || 'MANUEL'}-{new Date().getTime().toString().slice(-6)}</span></div>
                      <div>Généré le : <span className="text-slate-400 font-bold">{new Date().toLocaleString('fr-FR')}</span></div>
                      <div className="text-slate-400 font-bold">Signature Qualité requise : __________________</div>
                    </div>

                  </div>
                )}

              </div>

            </div>

          </div>
        )}

        {/* ========================================================================= */}
        {/* TAB 4: ADMIN PANEL - GESTION TECHNICIENS & HISTORIQUE DES CHANGEMENTS     */}
        {/* ========================================================================= */}
        {activeTab === 'admin-panel' && session?.role === 'admin' && (
          <div className="space-y-6 animate-fadeIn">
            <div className="bg-slate-900 border border-slate-800 p-6">
              <div className="flex items-center gap-3 mb-2 border-b border-slate-800 pb-4">
                <Shield className="h-6 w-6 text-amber-500" />
                <div>
                  <h2 className="text-lg font-black uppercase tracking-wider text-white font-display">
                    Panneau d'Administration Métrologie
                  </h2>
                  <p className="text-[11px] text-slate-400 font-mono uppercase tracking-wider">
                    Gestion exclusive des sessions techniciens et historique des modifications
                  </p>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 mt-6">
                
                {/* COLUMN 1: GESTION DES TECHNICIENS */}
                <div className="lg:col-span-5 bg-slate-950/40 border border-slate-900 p-5 space-y-5">
                  <div className="flex items-center gap-2 border-b border-slate-900 pb-2.5">
                    <User className="h-4.5 w-4.5 text-sky-400" />
                    <h3 className="text-xs font-black uppercase tracking-widest text-slate-200">
                      Gestion des Comptes Techniciens
                    </h3>
                  </div>

                  {/* Add Tech Form */}
                  <form onSubmit={handleAddTechnician} className="space-y-3.5 bg-slate-950 p-4 border border-slate-900">
                    <div className="text-[10px] font-bold text-sky-400 uppercase tracking-widest">
                      Ajouter un Nouveau Technicien
                    </div>
                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-widest mb-1">
                        Matricule du Technicien *
                      </label>
                      <input
                        type="text"
                        placeholder="Ex: MTR456"
                        value={newTechMatricule}
                        onChange={(e) => setNewTechMatricule(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                        required
                      />
                    </div>
                    <div>
                      <label className="block text-[9px] font-bold text-slate-500 uppercase tracking-widest mb-1">
                        Nom Complet *
                      </label>
                      <input
                        type="text"
                        placeholder="Ex: BEN MANSOUR Samir"
                        value={newTechName}
                        onChange={(e) => setNewTechName(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                        required
                      />
                    </div>
                    <button
                      type="submit"
                      className="w-full bg-sky-500 hover:bg-sky-400 text-slate-950 font-black py-2 px-3 text-[10px] uppercase tracking-wider flex items-center justify-center gap-1.5 transition-all cursor-pointer"
                    >
                      <PlusCircle className="h-3.5 w-3.5" />
                      <span>Enregistrer le Technicien</span>
                    </button>
                  </form>

                  {/* Registered Techs List */}
                  <div className="space-y-2">
                    <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                      Techniciens Enregistrés ({technicians.length})
                    </div>
                    {technicians.length === 0 ? (
                      <div className="text-center py-4 text-xs text-slate-600 font-mono border border-dashed border-slate-900">
                        Aucun technicien enregistré.
                      </div>
                    ) : (
                      <div className="space-y-1.5 max-h-60 overflow-y-auto pr-1">
                        {technicians.map((tech) => (
                          <div
                            key={tech.matricule}
                            className="bg-slate-950/80 border border-slate-900 px-3.5 py-2.5 flex items-center justify-between text-xs"
                          >
                            <div className="font-mono">
                              <span className="text-slate-500 font-bold mr-2">[{tech.matricule}]</span>
                              <span className="text-slate-200 font-bold">{tech.name}</span>
                            </div>
                            
                            <div className="flex items-center">
                              {techToDelete === tech.matricule ? (
                                <div className="flex items-center gap-1 bg-rose-950/60 border border-rose-900 p-1 font-mono animate-fadeIn">
                                  <span className="text-[9px] text-rose-300 font-bold px-1 uppercase">Sûr?</span>
                                  <button
                                    type="button"
                                    onClick={() => handleDeleteTechnician(tech.matricule)}
                                    className="text-[8px] bg-rose-600 hover:bg-rose-500 text-white font-bold px-1.5 py-0.5 uppercase cursor-pointer"
                                  >
                                    Oui
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => setTechToDelete(null)}
                                    className="text-[8px] bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold px-1.5 py-0.5 uppercase cursor-pointer"
                                  >
                                    Non
                                  </button>
                                </div>
                              ) : (
                                <button
                                  type="button"
                                  onClick={() => setTechToDelete(tech.matricule)}
                                  className="text-slate-500 hover:text-rose-400 p-1 hover:bg-rose-950/20 transition-all cursor-pointer"
                                  title="Supprimer ce technicien"
                                >
                                  <Trash2 className="h-3.5 w-3.5" />
                                </button>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                </div>

                {/* COLUMN 2: HISTORIQUE DES CHANGEMENTS */}
                <div className="lg:col-span-7 bg-slate-950/40 border border-slate-900 p-5 space-y-4">
                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 border-b border-slate-900 pb-2.5">
                    <div className="flex items-center gap-2">
                      <FileSpreadsheet className="h-4.5 w-4.5 text-amber-500" />
                      <h3 className="text-xs font-black uppercase tracking-widest text-slate-200">
                        Historique d'Audit & Changements
                      </h3>
                    </div>
                    
                    <div className="flex items-center gap-2 flex-wrap">
                      {logs.length > 0 && (
                        <button
                          type="button"
                          onClick={handleExportLogs}
                          className="text-[9px] bg-emerald-950/40 hover:bg-emerald-900/40 text-emerald-400 hover:text-emerald-300 border border-emerald-900/50 rounded-none px-2.5 py-1 font-mono uppercase font-bold transition-all cursor-pointer flex items-center gap-1"
                        >
                          <Download className="h-3 w-3" />
                          <span>Exporter (.txt)</span>
                        </button>
                      )}
                      
                      {logs.length > 0 && (
                        <div className="flex items-center">
                          {showClearConfirm ? (
                            <div className="flex items-center gap-1 bg-rose-950/60 border border-rose-900 p-1 font-mono animate-fadeIn">
                              <span className="text-[9px] text-rose-300 font-bold px-1 uppercase">Sûr?</span>
                              <button
                                type="button"
                                onClick={handleClearLogs}
                                className="text-[8px] bg-rose-600 hover:bg-rose-500 text-white font-bold px-1.5 py-0.5 uppercase cursor-pointer"
                              >
                                Oui
                              </button>
                              <button
                                type="button"
                                onClick={() => setShowClearConfirm(false)}
                                className="text-[8px] bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold px-1.5 py-0.5 uppercase cursor-pointer"
                              >
                                Non
                              </button>
                            </div>
                          ) : (
                            <button
                              type="button"
                              onClick={() => setShowClearConfirm(true)}
                              className="text-[9px] bg-rose-950/40 hover:bg-rose-900/40 text-rose-400 hover:text-rose-300 border border-rose-900/50 rounded-none px-2.5 py-1 font-mono uppercase font-bold transition-all cursor-pointer"
                            >
                              Vider l'historique
                            </button>
                          )}
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Search and Filters */}
                  <div className="relative">
                    <input
                      type="text"
                      placeholder="Filtrer l'historique par action, technicien, ou ID Master..."
                      value={logSearchQuery}
                      onChange={(e) => setLogSearchQuery(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-900 rounded-none px-3.5 py-2 text-xs text-slate-300 font-mono focus:outline-none focus:border-amber-500"
                    />
                  </div>

                  {/* Logs Table */}
                  <div className="border border-slate-900 bg-slate-950">
                    <div className="max-h-96 overflow-y-auto font-mono text-[11px]">
                      {filteredLogs.length === 0 ? (
                        <div className="text-center py-10 text-xs text-slate-600">
                          {logs.length === 0 ? "Aucun changement enregistré pour le moment." : "Aucun résultat pour cette recherche."}
                        </div>
                      ) : (
                        <div className="divide-y divide-slate-900">
                          {filteredLogs.map((log) => (
                            <div key={log.id} className="p-3.5 hover:bg-slate-900/30 transition-colors">
                              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1 mb-1.5">
                                <div className="flex items-center gap-2">
                                  <span className={`px-1.5 py-0.5 text-[8px] font-bold uppercase tracking-wider rounded-none ${
                                    log.action.includes("Création") ? "bg-emerald-950 text-emerald-400 border border-emerald-900" :
                                    log.action.includes("Suppression") ? "bg-rose-950 text-rose-400 border border-rose-900" :
                                    log.action.includes("Modification") ? "bg-sky-950 text-sky-400 border border-sky-900" :
                                    "bg-amber-950 text-amber-400 border border-amber-900"
                                  }`}>
                                    {log.action}
                                  </span>
                                  <span className="text-slate-500 text-[10px]">{log.timestamp}</span>
                                </div>
                                <div className="text-slate-400 text-[10px] font-bold">
                                  Par : <span className="text-sky-400">{log.user}</span>
                                </div>
                              </div>
                              <p className="text-slate-300 leading-relaxed break-words">{log.details}</p>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="text-[10px] text-slate-500 font-mono text-right uppercase">
                    Total: {logs.length} logs enregistrés localement
                  </div>
                </div>

              </div>
            </div>
          </div>
        )}

      </main>

      {/* Footer copyright */}
      <footer className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-16 pt-6 pb-8 border-t border-slate-900 flex flex-col sm:flex-row items-center justify-between gap-4 text-slate-500 text-xs">
        <div>
          <p className="font-bold text-slate-400">Master Management System</p>
          <p className="text-[11px] text-slate-600 mt-0.5">Logiciel interne d'assurance qualité et métrologie — ACTIA Tunisie.</p>
        </div>
        <div className="font-mono text-[10px] text-slate-400 uppercase tracking-widest bg-slate-900/40 border border-slate-900 px-4 py-2 hover:border-slate-800 transition-all">
          Created by <span className="text-sky-500 font-bold">Samir Ben Mansour</span>
        </div>
      </footer>

      {/* Custom Confirmation Modal for Deletion */}
      {masterToDelete && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fadeIn">
          <div className="w-full max-w-md bg-slate-900 border border-slate-800 p-6 shadow-2xl relative font-mono">
            {/* Header */}
            <div className="flex items-start justify-between gap-4 pb-4 border-b border-slate-800">
              <div className="flex items-center gap-2.5 text-rose-500">
                <AlertTriangle className="h-5 w-5 shrink-0 animate-pulse" />
                <h3 className="text-sm font-black uppercase tracking-wider font-display">Confirmation de suppression</h3>
              </div>
              <button 
                onClick={() => setMasterToDelete(null)}
                className="text-slate-500 hover:text-white transition-colors"
                title="Fermer"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            {/* Body */}
            <div className="py-4 space-y-4">
              <p className="text-xs text-slate-300 leading-relaxed font-sans">
                Êtes-vous sûr de vouloir supprimer définitivement cette fiche Master ? Cette opération est irréversible.
              </p>

              {/* Master item specifications */}
              <div className="bg-slate-950/60 p-3 border border-slate-800/80 space-y-2 text-[11px]">
                <div className="flex justify-between border-b border-slate-900 pb-1.5">
                  <span className="text-slate-500 uppercase font-bold text-[9px]">ID Master :</span>
                  <span className="text-sky-400 font-bold">{masterToDelete.idMaster}</span>
                </div>
                <div className="flex justify-between border-b border-slate-900 pb-1.5">
                  <span className="text-slate-500 uppercase font-bold text-[9px]">Testeur :</span>
                  <span className="text-slate-200 font-bold">{masterToDelete.testeur}</span>
                </div>
                <div className="flex justify-between border-b border-slate-900 pb-1.5">
                  <span className="text-slate-500 uppercase font-bold text-[9px]">Réf Produit :</span>
                  <span className="text-slate-200 font-bold">{masterToDelete.refProduitMaster || "-"}</span>
                </div>
                <div className="flex justify-between pb-0.5">
                  <span className="text-slate-500 uppercase font-bold text-[9px]">N° de série :</span>
                  <span className="text-slate-200 font-bold">{masterToDelete.numSerieProduitMaster || "-"}</span>
                </div>
              </div>
            </div>

            {/* Actions */}
            <div className="flex items-center justify-end gap-2.5 pt-4 border-t border-slate-800">
              <button
                type="button"
                onClick={() => setMasterToDelete(null)}
                className="bg-slate-950 hover:bg-slate-900 text-slate-400 font-bold px-4 py-2 text-[10px] uppercase tracking-wider border border-slate-800 transition-all cursor-pointer"
              >
                Annuler
              </button>
              <button
                type="button"
                onClick={confirmDeleteMaster}
                className="bg-rose-500 hover:bg-rose-400 text-slate-950 font-black px-4 py-2 text-[10px] uppercase tracking-wider border border-rose-400 shadow-[0_0_15px_rgba(244,63,94,0.2)] transition-all cursor-pointer"
              >
                Confirmer la suppression
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom Modal for Modification */}
      {masterToEdit && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fadeIn">
          <div className="w-full max-w-2xl bg-slate-900 border border-slate-800 p-6 shadow-2xl relative font-mono overflow-y-auto max-h-[90vh]">
            {/* Header */}
            <div className="flex items-start justify-between gap-4 pb-4 border-b border-slate-800">
              <div className="flex items-center gap-2.5 text-sky-400">
                <Edit3 className="h-5 w-5 shrink-0" />
                <h3 className="text-sm font-black uppercase tracking-wider font-display">Modification du Master</h3>
              </div>
              <button 
                onClick={() => setMasterToEdit(null)}
                className="text-slate-500 hover:text-white transition-colors"
                title="Fermer"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            {/* Body Form */}
            <div className="py-4 space-y-4">
              <p className="text-xs text-slate-400 leading-relaxed font-sans">
                Ajustez les paramètres métrologiques et les références du master <strong className="text-sky-400">{masterToEdit.idMaster}</strong>.
              </p>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs">
                {/* ID Master */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">ID Master *</label>
                  <input
                    type="text"
                    value={masterToEdit.idMaster}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, idMaster: e.target.value.toUpperCase() }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                    required
                  />
                </div>

                {/* Testeur de destination */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Testeur de destination *</label>
                  <input
                    type="text"
                    value={masterToEdit.testeur}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, testeur: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                    required
                  />
                </div>

                {/* Réf Produit */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Réf Produit Master</label>
                  <input
                    type="text"
                    value={masterToEdit.refProduitMaster || ''}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, refProduitMaster: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                  />
                </div>

                {/* N° Série Produit */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">N° Série Produit Master</label>
                  <input
                    type="text"
                    value={masterToEdit.numSerieProduitMaster || ''}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, numSerieProduitMaster: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                  />
                </div>

                {/* Réf Carte */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Réf Carte Master</label>
                  <input
                    type="text"
                    value={masterToEdit.refCarteMaster || ''}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, refCarteMaster: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                  />
                </div>

                {/* N° Série Carte */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">N° Série Carte Master</label>
                  <input
                    type="text"
                    value={masterToEdit.numSerieCarteMaster || ''}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, numSerieCarteMaster: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                  />
                </div>

                {/* Statut Selection */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Statut</label>
                  <select
                    value={masterToEdit.statut}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, statut: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                  >
                    <option value="Active">Active (Prête pour les bancs)</option>
                    <option value="Obsolète">Obsolète (Retirée de production)</option>
                    <option value="Endommagée">Endommagée (À réparer/Remplacer)</option>
                  </select>
                </div>

                {/* Verification Selection */}
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Vérification Métrologique</label>
                  <select
                    value={masterToEdit.verif}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, verif: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                  >
                    <option value="OK">OK (Mesures validées)</option>
                    <option value="KO">KO (Dérive ou défaut physique)</option>
                    <option value="FIN DE VIE">FIN DE VIE</option>
                  </select>
                </div>

                {/* Date d'enregistrement */}
                <div className="md:col-span-2">
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Date d'enregistrement</label>
                  <input
                    type="text"
                    value={masterToEdit.dateCreation}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, dateCreation: e.target.value }) : null)}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500"
                  />
                </div>

                {/* Commentaire 1 */}
                <div className="md:col-span-2">
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Commentaire principal</label>
                  <textarea
                    value={masterToEdit.commentaire1 || ''}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, commentaire1: e.target.value }) : null)}
                    rows={2}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500 resize-none animate-none"
                  />
                </div>

                {/* Commentaire 2 */}
                <div className="md:col-span-2">
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">Observations / Détails supplémentaires</label>
                  <textarea
                    value={masterToEdit.commentaire2 || ''}
                    onChange={(e) => setMasterToEdit(prev => prev ? ({ ...prev, commentaire2: e.target.value }) : null)}
                    rows={2}
                    className="w-full bg-slate-950 border border-slate-800 rounded-none px-3 py-2 text-xs text-slate-100 font-mono focus:outline-none focus:border-sky-500 resize-none animate-none"
                  />
                </div>
              </div>
            </div>

            {/* Actions */}
            <div className="flex items-center justify-end gap-2.5 pt-4 border-t border-slate-800">
              <button
                type="button"
                onClick={() => setMasterToEdit(null)}
                className="bg-slate-950 hover:bg-slate-900 text-slate-400 font-bold px-4 py-2 text-[10px] uppercase tracking-wider border border-slate-800 transition-all cursor-pointer"
              >
                Annuler
              </button>
              <button
                type="button"
                onClick={saveEditMaster}
                className="bg-sky-500 hover:bg-sky-400 text-slate-950 font-black px-4 py-2 text-[10px] uppercase tracking-wider border border-sky-400 shadow-[0_0_15px_rgba(14,165,233,0.2)] transition-all cursor-pointer"
              >
                Enregistrer les modifications
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Hidden high-fidelity printable template for professional PDF generation */}
      <div 
        id="pdf-report-template"
        className="absolute left-[-9999px] top-0 w-[800px] bg-white text-slate-800 p-10 font-sans border border-slate-200"
        style={{ display: "none" }}
      >
        {/* ACTIA Corporate Header */}
        <div className="flex items-center justify-between border-b-2 border-slate-900 pb-5 mb-6">
          <div className="flex items-center gap-3">
            {/* Inline high fidelity ACTIA Logo */}
            <svg viewBox="0 0 240 80" className="h-12 w-auto" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M 12 10 L 175 10 L 160 62 L -3 62 Z" fill="#5C5959" />
              <text x="20" y="48" fill="#FFFFFF" fontSize="36" fontWeight="900" fontFamily="system-ui, -apple-system, sans-serif" fontStyle="italic" letterSpacing="-1">ACTIA</text>
              <text x="142" y="28" fill="#FFFFFF" fontSize="10" fontWeight="bold" fontFamily="system-ui, -apple-system, sans-serif">®</text>
              <g transform="translate(142, 0)">
                <path d="M 38 10 L 51 10 L 46 24 L 33 24 Z" fill="#00965E" />
                <path d="M 53 10 L 66 10 L 61 24 L 48 24 Z" fill="#00965E" />
                <path d="M 68 10 L 81 10 L 76 24 L 63 24 Z" fill="#00965E" />
                <path d="M 83 10 L 96 10 L 91 24 L 78 24 Z" fill="#00965E" />
                <path d="M 33 27 L 46 27 L 41 41 L 28 41 Z" fill="#00965E" />
                <path d="M 48 27 L 61 27 L 56 41 L 43 41 Z" fill="#00965E" />
                <path d="M 63 27 L 76 27 L 71 41 L 58 41 Z" fill="#00965E" />
                <path d="M 78 27 L 91 27 L 86 41 L 73 41 Z" fill="#00965E" />
                <path d="M 28 44 L 41 44 L 36 58 L 23 58 Z" fill="#00965E" />
                <path d="M 43 44 L 56 44 L 51 58 L 38 58 Z" fill="#00965E" />
                <path d="M 58 44 L 71 44 L 66 58 L 53 58 Z" fill="#00965E" />
                <path d="M 73 44 L 86 44 L 81 58 L 68 58 Z" fill="#00965E" />
              </g>
              <text x="100" y="75" fill="#5C5959" fontSize="9" fontWeight="900" fontFamily="system-ui, -apple-system, sans-serif" letterSpacing="3">TUNISIE</text>
            </svg>
            <div className="border-l border-slate-300 pl-4">
              <h1 className="text-sm font-black tracking-wider text-slate-900 uppercase">DÉPARTEMENT QUALITÉ ET MÉTROLOGIE</h1>
              <p className="text-[10px] text-slate-500 font-mono">Assurance Qualité & Analyse Statistique (SPC)</p>
            </div>
          </div>
          <div className="text-right font-mono text-[10px] text-slate-500 space-y-0.5">
            <div>REF : ACTIA-SPC-CAP-{calcMasterId || "MANUEL"}</div>
            <div>Date : {new Date().toLocaleDateString('fr-FR')} {new Date().toLocaleTimeString('fr-FR')}</div>
            <div>Opérateur : {calcOperator || session?.name || "Technicien Qualité"}</div>
          </div>
        </div>

        {/* Document Title */}
        <div className="text-center mb-6">
          <h2 className="text-xl font-extrabold tracking-tight text-slate-900 uppercase">RAPPORT D'APTITUDE ET DE CAPABILITÉ PROCESSUS</h2>
          <p className="text-xs text-slate-500 font-mono mt-1">Conformément aux directives d'audit industriel IATF 16949 / ISO 9001</p>
        </div>

        {activeStats && (
          <div className="space-y-6">
            {/* Conformance Status Block */}
            {(() => {
              const reqThreshold = calcIsSc === 'OUI' ? 1.67 : 1.33;
              const reqWarning = calcIsSc === 'OUI' ? 1.33 : 1.0;
              const isCapable = activeStats.cpk >= reqThreshold;
              const isMarginal = activeStats.cpk >= reqWarning && activeStats.cpk < reqThreshold;
              
              let bgColor = "bg-emerald-50 border-emerald-300 text-emerald-800";
              let statusText = "CONFORME & QUALIFIÉ (PROCESSUS CAPABLE)";
              let descText = `Le moyen de mesure possède une capabilité de répétabilité optimale (Cpk ≥ ${reqThreshold.toFixed(2)}). Le banc est qualifié pour les contrôles de production sans restriction.`;
              
              if (isMarginal) {
                bgColor = "bg-amber-50 border-amber-300 text-amber-800";
                statusText = "MARGINAL (SURVEILLANCE ACCRUE REQUIS)";
                descText = `L'index de capabilité réelle est acceptable mais proche du seuil critique (${reqWarning.toFixed(2)} ≤ Cpk < ${reqThreshold.toFixed(2)}). Une surveillance périodique et un plan d'étalonnage sont préconisés.`;
              } else if (!isCapable && !isMarginal) {
                bgColor = "bg-rose-50 border-rose-300 text-rose-800";
                statusText = "NON CONFORME & REJETÉ (PROCESSUS INSUFFISANT)";
                descText = `Le processus présente une dispersion excessive ou un décentrage systématique (Cpk < ${reqWarning.toFixed(2)}). Le banc de test n'est pas apte à garantir la conformité métrologique.`;
              }

              return (
                <div className={`border p-4 rounded-none ${bgColor} flex items-center justify-between gap-4`}>
                  <div className="space-y-1">
                    <h3 className="text-sm font-black uppercase tracking-wider">{statusText}</h3>
                    <p className="text-[11px] leading-relaxed font-sans opacity-90">{descText}</p>
                  </div>
                  <div className="text-center border-l border-current pl-6 shrink-0">
                    <div className="text-3xl font-black font-mono tracking-tight">{activeStats.cpk.toFixed(3)}</div>
                    <div className="text-[9px] uppercase font-mono tracking-wider font-bold">Index Cpk Réel</div>
                  </div>
                </div>
              );
            })()}

            {/* General Information Tables */}
            <div className="grid grid-cols-2 gap-6 text-xs">
              {/* Left Column: Equipment Info */}
              <div className="border border-slate-200 p-4 space-y-3">
                <h3 className="font-extrabold text-slate-900 border-b border-slate-200 pb-1 uppercase tracking-wide text-[11px]">Caractéristiques du Moyen</h3>
                <table className="w-full">
                  <tbody className="divide-y divide-slate-100 font-sans">
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">ID Pièce Master :</td>
                      <td className="py-1.5 text-right font-mono font-bold text-slate-900">{calcMasterId || "Saisie Manuelle"}</td>
                    </tr>
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">Banc de Test (Testeur) :</td>
                      <td className="py-1.5 text-right font-bold text-slate-900">{calcTester || "-"}</td>
                    </tr>
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">Réf Produit Master :</td>
                      <td className="py-1.5 text-right text-slate-700">{masters.find(m => m.idMaster === calcMasterId)?.refProduitMaster || "SANS REF"}</td>
                    </tr>
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">N° de Série Master :</td>
                      <td className="py-1.5 text-right text-slate-700">{masters.find(m => m.idMaster === calcMasterId)?.numSerieProduitMaster || "SANS N/S"}</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Right Column: Spec Info */}
              <div className="border border-slate-200 p-4 space-y-3">
                <h3 className="font-extrabold text-slate-900 border-b border-slate-200 pb-1 uppercase tracking-wide text-[11px]">Tolérances & Spécifications</h3>
                <table className="w-full">
                  <tbody className="divide-y divide-slate-100 font-sans">
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">Grandeur / Paramètre :</td>
                      <td className="py-1.5 text-right font-bold text-slate-900">{calcCharacteristic}</td>
                    </tr>
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">Valeur Nominale :</td>
                      <td className="py-1.5 text-right font-mono font-bold text-slate-900">{calcNominal}</td>
                    </tr>
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">Limite Inférieure (LSL) :</td>
                      <td className="py-1.5 text-right font-mono text-slate-700">{calcLsl}</td>
                    </tr>
                    <tr>
                      <td className="py-1.5 text-slate-500 font-bold uppercase text-[10px]">Limite Supérieure (USL) :</td>
                      <td className="py-1.5 text-right font-mono text-slate-700">{calcUsl}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>

            {/* Core Statistics & Radar Chart Block */}
            <div className="grid grid-cols-5 gap-6">
              {/* Left Column: Synthèse des Indicateurs Statistiques */}
              <div className="col-span-3 border border-slate-200 p-4 space-y-3 flex flex-col justify-between">
                <div>
                  <h3 className="font-extrabold text-slate-900 border-b border-slate-200 pb-1.5 uppercase tracking-wide text-[11px]">Synthèse des Indicateurs Statistiques</h3>
                  <div className="grid grid-cols-2 gap-4 text-center font-mono py-1 text-xs">
                    <div className="border-r border-slate-100 pb-2">
                      <div className="text-[9px] uppercase text-slate-400 font-sans font-bold">Moyenne Processus (X-Bar)</div>
                      <div className="text-sm font-bold text-slate-900 mt-1">{activeStats.mean.toFixed(4)}</div>
                      <div className="text-[8px] text-slate-500 mt-0.5 font-sans">Cible : {calcNominal}</div>
                    </div>
                    <div className="pb-2">
                      <div className="text-[9px] uppercase text-slate-400 font-sans font-bold">Écart-Type (Sigma σ)</div>
                      <div className="text-sm font-bold text-slate-900 mt-1">{activeStats.stdDev.toFixed(5)}</div>
                      <div className="text-[8px] text-slate-500 mt-0.5 font-sans">Dispersion</div>
                    </div>
                    <div className="border-r border-slate-100 pt-2 border-t border-slate-100">
                      <div className="text-[9px] uppercase text-slate-400 font-sans font-bold">Capabilité Cp (Potentiel)</div>
                      <div className="text-sm font-bold text-slate-900 mt-1">{activeStats.cp.toFixed(3)}</div>
                      <div className="text-[8px] text-slate-500 mt-0.5 font-sans">Dispersion seule</div>
                    </div>
                    <div className="pt-2 border-t border-slate-100">
                      <div className="text-[9px] uppercase text-slate-400 font-sans font-bold">Capabilité Cpk (Réelle)</div>
                      <div className="text-sm font-black text-slate-900 mt-1">{activeStats.cpk.toFixed(3)}</div>
                      <div className="text-[8px] text-slate-500 mt-0.5 font-sans">Centrage ({activeStats.cpkType})</div>
                    </div>
                  </div>
                </div>
                <div className="grid grid-cols-3 gap-2 text-center font-mono py-2 border-t border-slate-100 text-[10px]">
                  <div>
                    <span className="text-slate-400 font-sans font-bold text-[9px] uppercase block">Mesures (n)</span>
                    <span className="font-bold text-slate-800">{activeStats.n} relevés</span>
                  </div>
                  <div className="border-x border-slate-100">
                    <span className="text-slate-400 font-sans font-bold text-[9px] uppercase block">Étendue (Range R)</span>
                    <span className="font-bold text-slate-800">{activeStats.range.toFixed(4)}</span>
                  </div>
                  <div>
                    <span className="text-slate-400 font-sans font-bold text-[9px] uppercase block">Variabilité (6σ)</span>
                    <span className="font-bold text-slate-800">{activeStats.repeatability.toFixed(4)}</span>
                  </div>
                </div>
              </div>

              {/* Right Column: D3 Radar Chart */}
              <div className="col-span-2 border border-slate-200 p-4 flex flex-col items-center justify-between text-center bg-slate-50">
                <h3 className="font-extrabold text-slate-900 border-b border-slate-200 pb-1.5 uppercase tracking-wide text-[11px] w-full text-center">
                  Qualité Métrologique (D3 Radar)
                </h3>
                <div className="w-full flex items-center justify-center pt-2">
                  <RadarChart
                    cp={activeStats.cp}
                    cpk={activeStats.cpk}
                    centering={(2 * (activeStats.mean - calcNominal)) / (calcUsl - calcLsl)}
                    sigma={activeStats.stdDev}
                    range={activeStats.range}
                    lsl={calcLsl}
                    usl={calcUsl}
                    nominal={calcNominal}
                    width={200}
                    height={180}
                    theme="light"
                  />
                </div>
              </div>
            </div>

            {/* SPC Trend analysis inside PDF */}
            {calcMasterId && (() => {
              const trend = getMasterTrend(calcMasterId);
              if (trend) {
                return (
                  <div className={`p-4 border font-mono space-y-2 ${trend.isWarning ? 'bg-amber-50 border-amber-300 text-amber-900' : 'bg-slate-50 border-slate-200 text-slate-800'}`}>
                    <div className="flex items-center justify-between border-b border-current/10 pb-1.5">
                      <span className="text-[10px] font-bold tracking-wider uppercase flex items-center gap-1">
                        ANALYSE DE TENDANCE HISTORIQUE & DÉRIVE SPC (3 CALCULS)
                      </span>
                      {trend.isWarning ? (
                        <span className="bg-amber-600 text-white font-bold px-1.5 py-0.5 text-[8px] uppercase font-black">
                          ⚠️ DÉRIVE DÉTECTÉE
                        </span>
                      ) : (
                        <span className="bg-emerald-600 text-white font-bold px-1.5 py-0.5 text-[8px] uppercase font-black">
                          STABLE (Sous contrôle)
                        </span>
                      )}
                    </div>
                    <div className="grid grid-cols-3 gap-4 text-xs font-sans">
                      <div>
                        <span className="text-[9px] text-slate-500 uppercase font-mono block">Moyenne SPC cumulée :</span>
                        <strong className="text-slate-900 font-mono">{trend.averageMean}</strong>
                      </div>
                      <div>
                        <span className="text-[9px] text-slate-500 uppercase font-mono block">Taux de dérive / Tolérance :</span>
                        <strong className={trend.isWarning ? "text-amber-700 font-mono font-bold" : "text-emerald-700 font-mono"}>{trend.driftPercentage}%</strong>
                      </div>
                      <div>
                        <span className="text-[9px] text-slate-500 uppercase font-mono block">Moyennes historiques :</span>
                        <span className="font-mono text-[10px] text-slate-700">{trend.means.join(" ➔ ")}</span>
                      </div>
                    </div>
                    <p className="text-[10px] text-slate-600 font-sans leading-relaxed pt-1 border-t border-current/10">
                      {trend.isWarning 
                        ? `Avertissement SPC : La moyenne glissante a dévié de plus de 5% de la bande d'intervalle. Biais métrologique persistant à étalonner.`
                        : `La moyenne cumulée reste centrée. Pas de biais systématique détecté sur le banc de test.`
                      }
                    </p>
                  </div>
                );
              }
              return null;
            })()}

            {/* High-Fidelity Light-Themed Chart */}
            <div className="border border-slate-200 p-4 space-y-2">
              <h3 className="font-extrabold text-slate-900 border-b border-slate-200 pb-1 uppercase tracking-wide text-[11px] flex justify-between">
                <span>Carte de Dispersion de Répétabilité (Graphique SPC)</span>
                <span className="text-[9px] text-slate-500 font-mono">Abscisse : Numéro de l'essai</span>
              </h3>
              <div className="w-full h-48 bg-slate-50 border border-slate-100 p-1">
                <svg className="w-full h-full" viewBox="0 0 600 180" preserveAspectRatio="none">
                  {/* Grid Lines */}
                  <line x1="40" y1="20" x2="580" y2="20" stroke="#e2e8f0" strokeDasharray="2" />
                  <line x1="40" y1="60" x2="580" y2="60" stroke="#e2e8f0" strokeDasharray="2" />
                  <line x1="40" y1="100" x2="580" y2="100" stroke="#e2e8f0" strokeDasharray="2" />
                  <line x1="40" y1="140" x2="580" y2="140" stroke="#e2e8f0" strokeDasharray="2" />

                  {(() => {
                    const arr = measurementsArray;
                    const n = arr.length;
                    
                    const maxLimit = Math.max(calcUsl, ...arr);
                    const minLimit = Math.min(calcLsl, ...arr);
                    const margin = (maxLimit - minLimit) * 0.15 || 1.0;
                    const topY = maxLimit + margin;
                    const bottomY = minLimit - margin;
                    const rangeY = topY - bottomY;

                    const getX = (index: number) => {
                      if (n <= 1) return 310;
                      return 40 + (index / (n - 1)) * (580 - 40);
                    };
                    const getY = (val: number) => {
                      return 20 + ((topY - val) / rangeY) * (150 - 20);
                    };

                    const points = arr.map((val, idx) => ({
                      x: getX(idx),
                      y: getY(val),
                      val,
                      idx
                    }));

                    const uslY = getY(calcUsl);
                    const lslY = getY(calcLsl);
                    const nomY = getY(calcNominal);
                    const meanY = getY(activeStats.mean);

                    return (
                      <>
                        {/* Safe tolerance zone fill */}
                        <rect x="40" y={Math.min(uslY, lslY)} width="540" height={Math.abs(lslY - uslY)} fill="#10b981" fillOpacity="0.02" />

                        {/* Reference line USL */}
                        <line x1="40" y1={uslY} x2="580" y2={uslY} stroke="#dc2626" strokeWidth="1.2" strokeDasharray="3 3" />
                        <text x="45" y={uslY - 4} fill="#dc2626" className="text-[8px] font-mono font-bold">USL: {calcUsl}</text>

                        {/* Reference line LSL */}
                        <line x1="40" y1={lslY} x2="580" y2={lslY} stroke="#dc2626" strokeWidth="1.2" strokeDasharray="3 3" />
                        <text x="45" y={lslY + 10} fill="#dc2626" className="text-[8px] font-mono font-bold">LSL: {calcLsl}</text>

                        {/* Nominal line */}
                        <line x1="40" y1={nomY} x2="580" y2={nomY} stroke="#2563eb" strokeWidth="0.8" strokeDasharray="4 4" />
                        <text x="540" y={nomY - 4} fill="#2563eb" className="text-[8px] font-mono">Nom: {calcNominal}</text>

                        {/* Process Mean line */}
                        <line x1="40" y1={meanY} x2="580" y2={meanY} stroke="#16a34a" strokeWidth="1.2" />
                        <text x="510" y={meanY + 10} fill="#16a34a" className="text-[8px] font-mono font-bold">Moyenne: {activeStats.mean.toFixed(3)}</text>

                        {/* Connection path */}
                        {points.length > 1 && (
                          <path
                            d={points.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x} ${p.y}`).join(' ')}
                            fill="none"
                            stroke="#0284c7"
                            strokeWidth="1.8"
                          />
                        )}

                        {/* Nodes */}
                        {points.map((p) => {
                          const isOut = p.val > calcUsl || p.val < calcLsl;
                          return (
                            <g key={p.idx}>
                              <circle
                                cx={p.x}
                                cy={p.y}
                                r="3.5"
                                fill={isOut ? "#dc2626" : "#0284c7"}
                                stroke="#ffffff"
                                strokeWidth="1"
                              />
                              <text x={p.x} y={p.y - 6} textAnchor="middle" fill="#64748b" className="text-[7px] font-mono">
                                #{p.idx + 1}
                              </text>
                            </g>
                          );
                        })}
                      </>
                    );
                  })()}
                </svg>
              </div>
            </div>

            {/* Structured Compact Values Grid */}
            <div className="border border-slate-200 p-4 space-y-2">
              <h3 className="font-extrabold text-slate-900 border-b border-slate-200 pb-1 uppercase tracking-wide text-[11px]">
                Relevés bruts des mesures individuelles (n = {activeStats.n})
              </h3>
              <div className="grid grid-cols-5 gap-2 text-center text-xs font-mono">
                {measurementsArray.map((v, idx) => {
                  const isOut = v > calcUsl || v < calcLsl;
                  return (
                    <div 
                      key={idx} 
                      className={`py-1 border ${isOut ? 'bg-rose-50 border-rose-300 text-rose-800 font-bold' : 'bg-slate-50 border-slate-100 text-slate-700'}`}
                    >
                      <span className="text-[8px] text-slate-400 font-sans block">Essai #{idx + 1}</span>
                      <span>{v.toFixed(4)}</span>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Quality Sign-off block */}
            <div className="pt-8 border-t border-slate-200 grid grid-cols-2 gap-12 text-xs">
              <div className="space-y-12">
                <div className="border-b border-slate-300 pb-1 font-bold text-slate-700 uppercase tracking-wider text-[10px]">SIGNATURE DE L'OPÉRATEUR</div>
                <div className="space-y-1">
                  <div className="font-bold text-slate-900">{calcOperator || session?.name || "Technicien Qualité"}</div>
                  <div className="text-[9px] text-slate-400 font-mono">Métrologie ACTIA — Date : {new Date().toLocaleDateString('fr-FR')}</div>
                </div>
              </div>
              <div className="space-y-12">
                <div className="border-b border-slate-300 pb-1 font-bold text-slate-700 uppercase tracking-wider text-[10px]">VISA RESPONSABLE QUALITÉ / LABORATOIRE</div>
                <div className="space-y-1">
                  <div className="text-slate-400 italic">Signature & Cachet autorisés</div>
                  <div className="text-[9px] text-slate-400 font-mono">ACTIA Tunisie — Département Métrologie</div>
                </div>
              </div>
            </div>

            {/* Corporate Footer disclaimer */}
            <div className="text-center text-[8px] text-slate-400 font-mono pt-4 border-t border-slate-100 uppercase tracking-wider">
              Document confidentiel à usage interne — Propriété de ACTIA Automotive Tunisie
            </div>
          </div>
        )}
      </div>

    </div>
  );
}
