import React, { useState, useMemo, useRef } from "react";
import { 
  Upload, 
  FileSpreadsheet, 
  FileText, 
  CheckCircle, 
  XCircle, 
  AlertCircle, 
  Play, 
  Settings, 
  Search, 
  Trash2, 
  Info, 
  Sliders, 
  Layers 
} from "lucide-react";
import { 
  ParsedLogFile, 
  CapabilityResult, 
  RepeatabilityResult, 
  AirBusEVResult, 
  DEFAULT_ACTIA_CONFIG, 
  ActipaCalculationConfig 
} from "./types";
import { parseActiaLog, aggregateTestSteps } from "./logParser";
import { 
  calculateCapability, 
  calculateRepeatability, 
  calculateAirbusEV, 
  assignGrrMetadata, 
  GrrFileAssignment 
} from "./calculations";
import { 
  exportCapabilityToExcel, 
  exportRepeatabilityToExcel, 
  exportAirbusToExcel 
} from "./excelExport";
import { MasterItem, UserSession } from "../types";

interface ActipaPanelProps {
  calcMasterId?: string;
  masters?: MasterItem[];
  session?: UserSession | null;
}

export default function ActipaPanel({ calcMasterId, masters = [], session }: ActipaPanelProps) {
  // Mode selection
  const [mode, setMode] = useState<'CAPABILITY' | 'REPEATABILITY'>(() => {
    const saved = localStorage.getItem("actipa_mode");
    return saved === 'REPEATABILITY' ? 'REPEATABILITY' : 'CAPABILITY';
  });
  
  // State for parsed files and config
  const [files, setFiles] = useState<ParsedLogFile[]>(() => {
    const saved = localStorage.getItem("actipa_files");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error("Error reading saved actipa files", e);
      }
    }
    return [];
  });
  const [config, setConfig] = useState<ActipaCalculationConfig>(() => {
    const saved = localStorage.getItem("actipa_config");
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        console.error("Error reading saved actipa config", e);
      }
    }
    return DEFAULT_ACTIA_CONFIG;
  });
  const [showConfig, setShowConfig] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");

  // Form Fields
  const selectedMaster = useMemo(() => {
    return masters.find(m => m.id === calcMasterId || m.idMaster === calcMasterId);
  }, [calcMasterId, masters]);

  const [productRef, setProductRef] = useState(() => {
    return localStorage.getItem("actipa_product_ref") || selectedMaster?.refProduitMaster || "";
  });
  const [operatorName, setOperatorName] = useState(() => {
    const saved = localStorage.getItem("actipa_operator_name");
    if (saved) return saved;
    if (session) {
      return `${session.name} (${session.matricule})`;
    }
    return "Technicien ACTIA";
  });
  const [isScMode, setIsScMode] = useState(() => {
    return localStorage.getItem("actipa_is_sc_mode") === "true";
  });

  // Sync operatorName when session changes
  React.useEffect(() => {
    if (session) {
      setOperatorName(`${session.name} (${session.matricule})`);
    }
  }, [session]);

  // List of unique testers for dropdown selection
  const uniqueTesters = useMemo(() => {
    const set = new Set(masters.map(m => m.testeur?.trim()).filter(Boolean));
    return Array.from(set).sort();
  }, [masters]);

  const [testerName, setTesterName] = useState(() => {
    return localStorage.getItem("actipa_tester_name") || selectedMaster?.testeur || "";
  });

  // Masters linked to the currently selected tester, excluding obsolete ones,
  // used to populate the dynamic "Référence Produit" dropdown.
  const masterOptionsForTester = useMemo(() => {
    if (!testerName) return [];
    return masters.filter(
      m =>
        m.testeur?.trim().toLowerCase() === testerName.trim().toLowerCase() &&
        m.statut?.trim().toLowerCase() !== "obsolète"
    );
  }, [testerName, masters]);

  // Unique product references (one per master line) for the selected tester.
  const productRefOptions = useMemo(() => {
    const set = new Set(masterOptionsForTester.map(m => m.refProduitMaster?.trim()).filter(Boolean));
    return Array.from(set).sort();
  }, [masterOptionsForTester]);

  // The master linked to the selection (tester + chosen product reference).
  const activeMaster = useMemo(() => {
    if (!testerName || !productRef) return null;
    return masterOptionsForTester.find(
      m => m.refProduitMaster?.trim().toLowerCase() === productRef.trim().toLowerCase()
    ) || null;
  }, [testerName, productRef, masterOptionsForTester]);

  // Results state
  const [analysisRan, setAnalysisRan] = useState(() => {
    return localStorage.getItem("actipa_analysis_ran") === "true";
  });
  const [capabilityResults, setCapabilityResults] = useState<CapabilityResult[]>(() => {
    const saved = localStorage.getItem("actipa_capability_results");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { console.error(e); }
    }
    return [];
  });
  const [repeatabilityResults, setRepeatabilityResults] = useState<RepeatabilityResult[]>(() => {
    const saved = localStorage.getItem("actipa_repeatability_results");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { console.error(e); }
    }
    return [];
  });
  const [airbusResults, setAirbusResults] = useState<AirBusEVResult[]>(() => {
    const saved = localStorage.getItem("actipa_airbus_results");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { console.error(e); }
    }
    return [];
  });
  const [rawMeasurementsMap, setRawMeasurementsMap] = useState<{ [testPointKey: string]: number[] }>(() => {
    const saved = localStorage.getItem("actipa_raw_measurements_map");
    if (saved) {
      try { return JSON.parse(saved); } catch (e) { console.error(e); }
    }
    return {};
  });

  const [viewTab, setViewTab] = useState<'REPORT' | 'TABLE'>(() => {
    return (localStorage.getItem("actipa_view_tab") as 'REPORT' | 'TABLE') || 'REPORT';
  });

  // Save Actipa states to localStorage on change
  React.useEffect(() => {
    localStorage.setItem("actipa_view_tab", viewTab);
    localStorage.setItem("actipa_mode", mode);
    localStorage.setItem("actipa_files", JSON.stringify(files));
    localStorage.setItem("actipa_config", JSON.stringify(config));
    localStorage.setItem("actipa_product_ref", productRef);
    localStorage.setItem("actipa_operator_name", operatorName);
    localStorage.setItem("actipa_is_sc_mode", String(isScMode));
    localStorage.setItem("actipa_tester_name", testerName);
    localStorage.setItem("actipa_analysis_ran", String(analysisRan));
    localStorage.setItem("actipa_capability_results", JSON.stringify(capabilityResults));
    localStorage.setItem("actipa_repeatability_results", JSON.stringify(repeatabilityResults));
    localStorage.setItem("actipa_airbus_results", JSON.stringify(airbusResults));
    localStorage.setItem("actipa_raw_measurements_map", JSON.stringify(rawMeasurementsMap));
  }, [
    mode,
    files,
    config,
    productRef,
    operatorName,
    isScMode,
    testerName,
    analysisRan,
    capabilityResults,
    repeatabilityResults,
    airbusResults,
    rawMeasurementsMap
  ]);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [fileWarning, setFileWarning] = useState<string>("");

  // Sync product reference & tester name when master is selected, and reset the
  // previously loaded files/logs so a stale list is never kept when switching pieces.
  const prevMasterIdRef = React.useRef<string | undefined>(undefined);
  React.useEffect(() => {
    if (selectedMaster) {
      if (selectedMaster.refProduitMaster) {
        setProductRef(selectedMaster.refProduitMaster);
      }
      if (selectedMaster.testeur) {
        setTesterName(selectedMaster.testeur);
      }
    }
    // When the selected master actually changes (e.g. clicking CALCULER on another
    // piece), clear any imported files from the previous testeur.
    if (selectedMaster?.id !== prevMasterIdRef.current) {
      prevMasterIdRef.current = selectedMaster?.id;
      clearFiles();
    }
  }, [selectedMaster]);

  // Handle files
  const handleFilesAdded = async (fileList: FileList) => {
    const selectedCount = fileList.length;
    if (selectedCount === 0) return;

    // Strict rule: a minimum of 10 files is required to run a calculation.
    if (files.length + selectedCount < 10) {
      setFileWarning(`Minimum 10 fichiers requis pour lancer le calcul. Actuellement sélectionnés : ${files.length + selectedCount}.`);
      return;
    }
    setFileWarning("");

    const newFiles: ParsedLogFile[] = [];
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
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
        const parsed = parseActiaLog(text, file.name);
        newFiles.push(parsed);
      } catch (err) {
        console.error(`Erreur lecture fichier ${file.name}:`, err);
      }
    }

    setFiles(prev => {
      const updated = [...prev, ...newFiles];
      // Try to auto-set productRef from first file if blank
      if (updated.length > 0 && !productRef) {
        const firstRef = updated.find(f => f.productRef)?.productRef;
        if (firstRef) setProductRef(firstRef);
      }
      return updated;
    });
    setAnalysisRan(false);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files) {
      await handleFilesAdded(e.dataTransfer.files);
    }
  };

  const clearFiles = () => {
    setFiles([]);
    setAnalysisRan(false);
    setCapabilityResults([]);
    setRepeatabilityResults([]);
    setAirbusResults([]);
    setRawMeasurementsMap({});
    setFileWarning("");
  };

  // Changing the tester/banc must reset the previously loaded files (and results)
  // so a stale file list from another testeur is never reused.
  const handleTesterChange = (value: string) => {
    setTesterName(value);
    const validMasters = masters.filter(
      m => m.testeur?.trim().toLowerCase() === value.trim().toLowerCase() &&
        m.statut?.trim().toLowerCase() !== "obsolète"
    );
    setProductRef(validMasters[0]?.refProduitMaster || "");
    clearFiles();
  };

  // Run calculation analysis
  const runAnalysis = () => {
    if (files.length === 0) return;

    // Sort files alphabetically to match assignment indexing perfectly
    const sortedFiles = [...files].sort((a, b) => a.filename.localeCompare(b.filename));

    // Aggregate values
    const steps = aggregateTestSteps(sortedFiles);
    
    // Store raw values map for exports
    const rawMap: { [key: string]: number[] } = {};
    for (const step of steps) {
      rawMap[`${step.testPointId}::${step.testStepName}`] = step.values;
    }
    setRawMeasurementsMap(rawMap);

    if (mode === 'CAPABILITY') {
      const results = steps.map(step => calculateCapability(step, config, isScMode));
      setCapabilityResults(results);
    } else if (mode === 'REPEATABILITY') {
      const assignments = sortedFiles.map((f, idx) => assignGrrMetadata(f.filename, idx, sortedFiles.length));
      
      const results = steps.map(step => calculateRepeatability(step, assignments, config, isScMode));
      setRepeatabilityResults(results);
    } else {
      const results = steps.map(step => calculateAirbusEV(step, config, isScMode));
      setAirbusResults(results);
    }

    setAnalysisRan(true);
  };

  // Download Excel Report
  const triggerExcelDownload = async () => {
    if (!analysisRan) return;

    const overallTestName = files[0]?.testName;

    if (mode === 'CAPABILITY') {
      await exportCapabilityToExcel(capabilityResults, rawMeasurementsMap, productRef, testerName, operatorName, overallTestName);
    } else if (mode === 'REPEATABILITY') {
      await exportRepeatabilityToExcel(repeatabilityResults, rawMeasurementsMap, productRef, testerName, operatorName, files, overallTestName);
    } else {
      await exportAirbusToExcel(airbusResults, rawMeasurementsMap, productRef, testerName, operatorName);
    }
  };

  // Get GRR Metadata Assignments for Display
  const grrAssignments = useMemo(() => {
    if (mode !== 'REPEATABILITY' || files.length === 0) return [];
    const sorted = [...files].sort((a, b) => a.filename.localeCompare(b.filename));
    return sorted.map((f, idx) => {
      const meta = assignGrrMetadata(f.filename, idx, sorted.length);
      return {
        filename: f.filename,
        operator: `Opérateur ${meta.operatorIndex + 1}`,
        part: `Pièce ${meta.partIndex + 1}`,
        trial: `Essai ${meta.trialIndex + 1}`
      };
    });
  }, [mode, files]);

  // Filtered Results to display
  const filteredCapability = useMemo(() => {
    return capabilityResults.filter(r => 
      r.testPointId.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.testStepName.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [capabilityResults, searchQuery]);

  const filteredRepeatability = useMemo(() => {
    return repeatabilityResults.filter(r => 
      r.testPointId.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.testStepName.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [repeatabilityResults, searchQuery]);

  const filteredAirbus = useMemo(() => {
    return airbusResults.filter(r => 
      r.testPointId.toLowerCase().includes(searchQuery.toLowerCase()) ||
      r.testStepName.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [airbusResults, searchQuery]);

  // Counts of OK / NOK
  const statsSummary = useMemo(() => {
    let ok = 0;
    let nok = 0;
    let nd = 0;

    if (mode === 'CAPABILITY' && analysisRan) {
      capabilityResults.forEach(r => {
        if (r.status === 'OK') ok++;
        else if (r.status === 'NOK') nok++;
        else nd++;
      });
    } else if (mode === 'REPEATABILITY' && analysisRan) {
      repeatabilityResults.forEach(r => {
        if (r.status === 'OK') ok++;
        else if (r.status === 'NOK') nok++;
        else nd++;
      });
    } else if (mode === 'AIRBUS_EV' && analysisRan) {
      airbusResults.forEach(r => {
        if (r.status === 'OK') ok++;
        else if (r.status === 'NOK') nok++;
        else nd++;
      });
    }

    return { ok, nok, nd, total: ok + nok + nd };
  }, [mode, analysisRan, capabilityResults, repeatabilityResults, airbusResults]);

  return (
    <div className="space-y-6">
      
      {/* HEADER SECTION - Mode Selector & Description */}
      <div className="bg-slate-900 border border-slate-800 p-6 flex flex-col md:flex-row md:items-center justify-between gap-6 shadow-xl">
        <div className="space-y-2">
          <div className="flex items-center gap-3">
            <Layers className="h-6 w-6 text-rose-500 animate-pulse" />
            <h2 className="font-display font-black text-xl text-white uppercase tracking-tight">ActIPA Metrology Suite v3.1</h2>
          </div>
          <p className="text-xs text-slate-400 max-w-xl">
            Calculateur de capabilité SPC (Cp, Cpk), répétabilité GRR / MSA selon référentiels industriels.
          </p>
        </div>
        
        {/* Tab Buttons */}
        <div className="flex p-1 bg-slate-950 border border-slate-800 rounded-none self-start md:self-center shrink-0">
          <button
            onClick={() => { setMode('CAPABILITY'); setAnalysisRan(false); }}
            className={`px-4 py-2 text-xs font-mono font-bold tracking-wider uppercase transition-all rounded-none ${
              mode === 'CAPABILITY' 
                ? 'bg-rose-600 text-white' 
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Capabilité
          </button>
          <button
            onClick={() => { setMode('REPEATABILITY'); setAnalysisRan(false); }}
            className={`px-4 py-2 text-xs font-mono font-bold tracking-wider uppercase transition-all rounded-none ${
              mode === 'REPEATABILITY' 
                ? 'bg-rose-600 text-white' 
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Répétabilité GRR
          </button>
        </div>
      </div>

      {/* TWO COLUMNS INPUTS & UPLOAD */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
        
        {/* Column 1: Config and Metadata Parameters */}
        <div className="lg:col-span-5 bg-slate-900/40 border border-slate-900 p-5 space-y-6 flex flex-col justify-between">
          <div className="space-y-5">
            <div className="flex items-center justify-between border-b border-slate-900 pb-3">
              <span className="text-xs font-mono font-black text-rose-400 uppercase tracking-widest flex items-center gap-2">
                <Sliders className="h-4 w-4" /> 1. Paramètres d'Analyse
              </span>
              <button 
                onClick={() => setShowConfig(!showConfig)}
                className="text-slate-400 hover:text-white text-xs font-mono flex items-center gap-1 border border-slate-800 hover:border-slate-700 px-2 py-1 bg-slate-950 transition-colors"
              >
                <Settings className="h-3.5 w-3.5" /> Coefficients
              </button>
            </div>

            {/* Threshold Configurations panel */}
            {showConfig && (
              <div className="p-4 bg-slate-950 border border-slate-800 space-y-3 font-mono text-[11px] text-slate-300">
                <div className="font-bold text-white text-xs border-b border-slate-900 pb-1 mb-2 uppercase">Configuration des Coefficients</div>
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">Lim Cpk Normal</label>
                    <input 
                      type="number" step="0.01" value={config.cpkLimNormal}
                      onChange={e => setConfig(prev => ({...prev, cpkLimNormal: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">Lim Cpk CS</label>
                    <input 
                      type="number" step="0.01" value={config.cpkLimSc}
                      onChange={e => setConfig(prev => ({...prev, cpkLimSc: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">Lim %GRR Normal</label>
                    <input 
                      type="number" step="0.01" value={config.msaLimNormal}
                      onChange={e => setConfig(prev => ({...prev, msaLimNormal: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">Lim %GRR CS</label>
                    <input 
                      type="number" step="0.01" value={config.msaLimSc}
                      onChange={e => setConfig(prev => ({...prev, msaLimSc: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">Lim NDC</label>
                    <input 
                      type="number" step="0.1" value={config.ndcLim}
                      onChange={e => setConfig(prev => ({...prev, ndcLim: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">K1 (m=3)</label>
                    <input 
                      type="number" step="0.0001" value={config.k1_m3}
                      onChange={e => setConfig(prev => ({...prev, k1_m3: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">K2 (op=2)</label>
                    <input 
                      type="number" step="0.0001" value={config.k2}
                      onChange={e => setConfig(prev => ({...prev, k2: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                  <div>
                    <label className="block text-slate-500 uppercase text-[9px] font-bold">K1 Airbus</label>
                    <input 
                      type="number" step="0.0001" value={config.k1_airbus}
                      onChange={e => setConfig(prev => ({...prev, k1_airbus: parseFloat(e.target.value)}))}
                      className="w-full bg-slate-900 border border-slate-800 p-1.5 mt-0.5 text-white"
                    />
                  </div>
                </div>
              </div>
            )}

            {/* Metrology details */}
            <div className="space-y-4 font-mono text-xs">
              <div>
                <label className="block text-slate-500 uppercase text-[10px] font-bold tracking-wider">Sélectionner un Testeur / Banc</label>
                <select 
                  value={uniqueTesters.includes(testerName) ? testerName : ""}
                  onChange={e => {
                    handleTesterChange(e.target.value);
                  }}
                  className="w-full bg-slate-950 border border-slate-850 p-2.5 mt-1 text-white focus:outline-none focus:border-rose-500 font-mono"
                >
                  <option value="">-- Sélectionner un testeur (ou saisir ci-dessous) --</option>
                  {uniqueTesters.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
                
                <input 
                  type="text" 
                  value={testerName} 
                  onChange={e => {
                    handleTesterChange(e.target.value);
                  }}
                  placeholder="Nom du Testeur / Banc de Test"
                  className="w-full bg-slate-950 border border-slate-850 p-2.5 mt-1.5 text-white focus:outline-none focus:border-rose-500"
                />
              </div>

              <div>
                <label className="block text-slate-500 uppercase text-[10px] font-bold tracking-wider">Référence Produit</label>
                <select
                  value={productRefOptions.includes(productRef) ? productRef : ""}
                  onChange={e => {
                    setProductRef(e.target.value);
                    clearFiles();
                  }}
                  disabled={!testerName || productRefOptions.length === 0}
                  className="w-full bg-slate-950 border border-slate-850 p-2.5 mt-1 text-white focus:outline-none focus:border-rose-500 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <option value="">
                    {!testerName
                      ? "-- Sélectionnez d'abord un banc --"
                      : productRefOptions.length === 0
                        ? "-- Aucune référence active pour ce banc --"
                        : "-- Choisir une référence produit --"}
                  </option>
                  {productRefOptions.map(ref => (
                    <option key={ref} value={ref}>{ref}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-slate-500 uppercase text-[10px] font-bold tracking-wider">Nom Opérateur</label>
                <input 
                  type="text" 
                  value={operatorName} 
                  onChange={e => setOperatorName(e.target.value)}
                  placeholder="Ex: Technicien ACTIA"
                  className="w-full bg-slate-950 border border-slate-850 p-2.5 mt-1 text-white focus:outline-none focus:border-rose-500"
                />
              </div>

              <div className="flex items-center gap-3 bg-slate-950/50 p-3 border border-slate-900">
                <input 
                  type="checkbox" 
                  id="chk_sc_mode"
                  checked={isScMode} 
                  onChange={e => setIsScMode(e.target.checked)}
                  className="h-4 w-4 bg-slate-900 border-slate-850 text-rose-500 focus:ring-0 cursor-pointer"
                />
                <label htmlFor="chk_sc_mode" className="text-xs text-slate-300 font-bold uppercase tracking-wider cursor-pointer">
                  Caractéristique Spéciale (CS / SC)
                </label>
              </div>

              {activeMaster && (
                <div className="p-3 bg-slate-950 border border-slate-850 border-l-4 border-l-rose-500 text-[11px] leading-relaxed text-slate-400">
                  <div className="font-bold text-slate-200 uppercase text-[9px] tracking-wider mb-1">Master Lié</div>
                  Banc : <span className="text-white font-bold">{activeMaster.testeur}</span> | REF : <span className="text-white font-bold">{activeMaster.refProduitMaster || "N/A"}</span> | Statut : <span className="text-white font-bold">{activeMaster.statut}</span>
                </div>
              )}
            </div>
          </div>

          {/* Bottom Run Action Button */}
          <div className="border-t border-slate-900 pt-4 mt-6">
            <button
              onClick={runAnalysis}
              disabled={files.length === 0}
              className={`w-full font-bold px-6 py-4 text-xs font-mono uppercase tracking-widest flex items-center justify-center gap-2 border transition-all ${
                files.length > 0
                  ? 'bg-rose-600 hover:bg-rose-500 text-white border-rose-500 shadow-[0_0_15px_rgba(225,29,72,0.15)] cursor-pointer'
                  : 'bg-slate-950 text-slate-600 border-slate-900 cursor-not-allowed'
              }`}
            >
              <Play className="h-4 w-4" /> Calculer les indicateurs ({files.length} fichiers)
            </button>
          </div>
        </div>

        {/* Column 2: Drag & Drop Multi Upload */}
        <div className="lg:col-span-7 flex flex-col gap-4">
          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`flex-1 border-2 border-dashed p-8 flex flex-col items-center justify-center text-center cursor-pointer transition-all min-h-[220px] ${
              isDragging 
                ? 'border-rose-500 bg-rose-950/10' 
                : 'border-slate-800 hover:border-slate-700 bg-slate-900/20'
            }`}
          >
            <input
              type="file"
              ref={fileInputRef}
              onChange={(e) => {
                if (e.target.files) handleFilesAdded(e.target.files);
              }}
              multiple
              accept=".log,.txt"
              className="hidden"
            />
            <Upload className="h-10 w-10 text-slate-500 mb-3 animate-bounce" />
            <div className="font-mono text-xs font-bold text-slate-200 uppercase tracking-widest">
              Glissez-déposez des logs ACTIA ou cliquez
            </div>
            <p className="text-[11px] text-slate-500 font-mono mt-2 uppercase">
              Fichiers .TXT ou .LOG (L301 / L302 / L303)
            </p>
          </div>

          {fileWarning && (
            <div className="bg-rose-950/40 border border-rose-700 text-rose-300 text-[11px] font-mono px-3 py-2 uppercase tracking-wide">
              {fileWarning}
            </div>
          )}

          {/* Files loaded info / Summary box */}
          {files.length > 0 && (
            <div className="bg-slate-900/60 border border-slate-900 p-4 space-y-3 font-mono">
              <div className="flex items-center justify-between border-b border-slate-950 pb-2">
                <span className="text-xs font-bold text-slate-300 uppercase tracking-wider flex items-center gap-1.5">
                  <FileText className="h-4 w-4 text-rose-500" /> Fichiers Chargés ({files.length})
                </span>
                <button 
                  onClick={clearFiles}
                  className="text-[10px] text-rose-400 hover:text-rose-300 uppercase font-bold flex items-center gap-1 hover:underline bg-transparent"
                >
                  <Trash2 className="h-3 w-3" /> Vider
                </button>
              </div>

              {/* Scrollable list of files with status badges */}
              <div className="max-h-[140px] overflow-y-auto divide-y divide-slate-950 text-[10px] pr-2 scrollbar-thin">
                {files.map((file, idx) => (
                  <div key={idx} className="py-2 flex items-center justify-between gap-3 text-slate-400">
                    <span className="truncate max-w-[240px]">{file.filename}</span>
                    <div className="flex items-center gap-2">
                      <span className="text-[9px] bg-slate-950 px-2 py-0.5 text-slate-500 font-semibold">{file.measurements.length} pts</span>
                      <span className={`px-2 py-0.5 text-[9px] font-black tracking-wider ${
                        file.status === 'PASS' 
                          ? 'bg-emerald-950 text-emerald-400 border border-emerald-900' 
                          : 'bg-rose-950 text-rose-400 border border-rose-900'
                      }`}>
                        {file.status}
                      </span>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Répétabilité files assignment info */}
      {mode === 'REPEATABILITY' && grrAssignments.length > 0 && (
        <div className="bg-slate-900/20 border border-slate-900 p-4 font-mono text-[10px] space-y-2">
          <div className="text-rose-400 font-bold uppercase tracking-wider flex items-center gap-1.5">
            <Info className="h-3.5 w-3.5" /> Attribution Automatique pour le calcul GRR (Opérateur / Pièce / Essai) :
          </div>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-2 max-h-[120px] overflow-y-auto pr-2">
            {grrAssignments.map((a, idx) => (
              <div key={idx} className="bg-slate-950 border border-slate-850 p-2 space-y-0.5">
                <div className="text-slate-500 truncate font-bold">{a.filename}</div>
                <div className="text-slate-300 font-bold">{a.operator}</div>
                <div className="text-slate-400">{a.part} • {a.trial}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ANALYSIS SUMMARY RESULTS */}
      {analysisRan && (
        <div className="space-y-6">
          
          {/* Quick numbers widget */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 bg-slate-900 border border-slate-800 p-5 font-mono uppercase">
            <div className="space-y-1">
              <span className="text-[10px] text-slate-500 font-bold tracking-wider">Analysés</span>
              <div className="text-2xl font-black text-white">{statsSummary.total}</div>
            </div>
            <div className="space-y-1">
              <span className="text-[10px] text-slate-500 font-bold tracking-wider">Statut Conforme (OK)</span>
              <div className="text-2xl font-black text-emerald-400">{statsSummary.ok}</div>
            </div>
            <div className="space-y-1">
              <span className="text-[10px] text-slate-500 font-bold tracking-wider">Non Conforme (NOK)</span>
              <div className="text-2xl font-black text-rose-400">{statsSummary.nok}</div>
            </div>
            <div className="space-y-1">
              <span className="text-[10px] text-slate-500 font-bold tracking-wider">Non Déterminé (ND)</span>
              <div className="text-2xl font-black text-slate-400">{statsSummary.nd}</div>
            </div>
          </div>

          {/* MAIN RESULTS TABLE AND PREVIEW */}
          <div className="space-y-4">
            {/* Tab Switches */}
            <div className="flex border-b border-slate-850 gap-1 font-mono text-xs">
              <button
                onClick={() => setViewTab('REPORT')}
                className={`px-5 py-3 font-bold tracking-wider uppercase transition-all duration-150 border-t-2 ${
                  viewTab === 'REPORT'
                    ? 'bg-slate-900 text-rose-400 border-t-rose-500 border-x border-x-slate-800'
                    : 'text-slate-500 hover:text-slate-300 border-t-transparent hover:bg-slate-900/40 cursor-pointer'
                }`}
              >
                📋 Aperçu du Rapport ({mode === 'CAPABILITY' ? 'AT242' : 'AT243'})
              </button>
              <button
                onClick={() => setViewTab('TABLE')}
                className={`px-5 py-3 font-bold tracking-wider uppercase transition-all duration-150 border-t-2 ${
                  viewTab === 'TABLE'
                    ? 'bg-slate-900 text-rose-400 border-t-rose-500 border-x border-x-slate-800'
                    : 'text-slate-500 hover:text-slate-300 border-t-transparent hover:bg-slate-900/40 cursor-pointer'
                }`}
              >
                📊 Tableau des Données
              </button>
            </div>

            {viewTab === 'REPORT' ? (
              /* Beautiful Report Sheet Preview */
              <div className="space-y-6">
                <div className="bg-white text-slate-900 p-6 sm:p-8 border border-slate-300 shadow-2xl rounded-none max-w-4xl mx-auto font-sans leading-normal">
                  {/* Top Header Block */}
                  <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pb-4 border-b border-slate-300">
                    <div className="flex items-center gap-3">
                      <img 
                        src="/actipa-assets/LOGO_ACTIA.png" 
                        onError={(e) => {
                          const target = e.target as HTMLImageElement;
                          if (target.src.endsWith(".png")) {
                            target.src = "/actipa-assets/LOGO_ACTIA.jpg";
                          } else {
                            target.style.display = "none";
                          }
                        }}
                        className="h-10 w-auto object-contain" 
                        referrerPolicy="no-referrer"
                        alt="ACTIA Logo" 
                      />
                      <span className="font-mono text-xs font-bold tracking-wider text-slate-400">ACTIA AUTOMOTIVE</span>
                    </div>
                    <div className="text-center font-bold text-[10px] text-red-600 border border-red-300 px-3 py-1 uppercase tracking-widest bg-red-50">
                      [PROJECT] RESTRICTED
                    </div>
                    <div className="text-right font-mono text-xs text-slate-500 font-bold">
                      {mode === 'CAPABILITY' ? 'AT242 - V3.0' : mode === 'REPEATABILITY' ? 'AT243 - V3.0' : 'AT-AIRBUS - V1.0'}
                    </div>
                  </div>

                  {/* Title Block */}
                  <div className="text-center py-6">
                    <h1 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-800 uppercase font-serif">
                      {mode === 'CAPABILITY' && "Capability ProcessAnalysis Report"}
                      {mode === 'REPEATABILITY' && "Measurement SystemAnalysis Report"}
                      {mode === 'AIRBUS_EV' && "Airbus EV% Metrology Report"}
                    </h1>
                    <p className="text-[10px] text-slate-400 font-mono uppercase mt-1">Saisie Automatique via ActIPA Suite v3</p>
                  </div>

                  {/* Metadata Table (Excel Style Grid) */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 border border-slate-200 p-4 bg-slate-50">
                    <div className="space-y-2 text-xs">
                      <div className="flex border-b border-slate-200 pb-1.5">
                        <span className="w-1/2 font-semibold text-slate-500">Product Name :</span>
                        <span className="w-1/2 font-bold text-slate-900">{testerName || "Banc"} - {productRef || "PRODUIT"}</span>
                      </div>
                      <div className="flex border-b border-slate-200 pb-1.5">
                        <span className="w-1/2 font-semibold text-slate-500">Product Reference :</span>
                        <span className="w-1/2 font-bold text-slate-900">{productRef || "N/A"}</span>
                      </div>
                      <div className="flex border-b border-slate-200 pb-1.5">
                        <span className="w-1/2 font-semibold text-slate-500">HW / SW Release :</span>
                        <span className="w-1/2 text-slate-900">V1.0</span>
                      </div>
                    </div>
                    <div className="space-y-2 text-xs">
                      <div className="flex border-b border-slate-200 pb-1.5">
                        <span className="w-1/2 font-semibold text-slate-500">EQUIPMENT REF. :</span>
                        <span className="w-1/2 font-bold text-slate-900">{testerName || "N/A"}</span>
                      </div>
                      <div className="flex border-b border-slate-200 pb-1.5">
                        <span className="w-1/2 font-semibold text-slate-500">EQUIPMENT DESIGNATION :</span>
                        <span className="w-1/2 text-slate-900">{testerName ? `${testerName} Bench` : "Functional Test Bench"}</span>
                      </div>
                      <div className="flex border-b border-slate-200 pb-1.5">
                        <span className="w-1/2 font-semibold text-slate-500">Analyse Type :</span>
                        <span className="w-1/2 font-bold uppercase text-rose-600">{mode}</span>
                      </div>
                    </div>
                  </div>

                  {/* REPORT GENERAL STATUS */}
                  <div className="my-6 border border-slate-200 p-4 text-center bg-slate-50">
                    <h2 className="text-[10px] font-mono font-bold text-slate-500 uppercase tracking-widest mb-2">REPORT GENERAL STATUS</h2>
                    {statsSummary.nok === 0 ? (
                      <div className="inline-flex items-center gap-2 px-6 py-2 bg-emerald-100 text-emerald-800 border border-emerald-300 rounded-none text-sm font-bold uppercase tracking-wider shadow-sm">
                        <CheckCircle className="h-4 w-4 text-emerald-600" /> TOUT CONFORME (OK)
                      </div>
                    ) : (
                      <div className="inline-flex items-center gap-2 px-6 py-2 bg-rose-100 text-rose-800 border border-rose-300 rounded-none text-sm font-bold uppercase tracking-wider shadow-sm">
                        <XCircle className="h-4 w-4 text-rose-600" /> PRÉSENCE DE NON-CONFORMITÉS (NOK)
                      </div>
                    )}
                    <p className="text-[10px] text-slate-400 mt-1.5 italic font-mono">
                      {statsSummary.ok} / {statsSummary.total} points conformes aux exigences métrologiques {isScMode ? "(Seuil SC)" : "(Seuil Standard)"}
                    </p>
                  </div>

                  {/* EVOLUTION TRACKING */}
                  <div className="my-6">
                    <h3 className="text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-2">Doc Evolution Tracking</h3>
                    <div className="border border-slate-200 overflow-hidden text-xs">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-100 border-b border-slate-200 text-[10px] font-bold uppercase text-slate-500">
                            <th className="px-3 py-2 border-r border-slate-200 w-24">Doc. Index</th>
                            <th className="px-3 py-2 border-r border-slate-200 w-32">Date</th>
                            <th className="px-3 py-2">Evolution description</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200 text-slate-700 font-mono text-[11px]">
                          <tr>
                            <td className="px-3 py-2 border-r border-slate-200">Index 01</td>
                            <td className="px-3 py-2 border-r border-slate-200">{new Date().toLocaleDateString('fr-FR')}</td>
                            <td className="px-3 py-2 font-sans text-slate-600">First automatic analysis generation on ActIPA Metrology Suite.</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* APPROVALS */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 border border-slate-200 p-4 mt-6 bg-slate-50 text-xs">
                    <div>
                      <div className="font-bold text-slate-500 uppercase text-[9px] tracking-wider border-b border-slate-200 pb-1 mb-2">APPROVAL AUTHOR</div>
                      <div className="space-y-1">
                        <div>Nom: <span className="font-bold text-slate-850">{operatorName || "Technicien ACTIA"}</span></div>
                        <div>Date: <span className="font-mono text-slate-600">{new Date().toLocaleDateString('fr-FR')}</span></div>
                        <div>Signature: <span className="font-mono text-emerald-600 font-bold">[SIGNATURE ÉLECTRONIQUE]</span></div>
                      </div>
                    </div>
                    <div>
                      <div className="font-bold text-slate-500 uppercase text-[9px] tracking-wider border-b border-slate-200 pb-1 mb-2">ACTIA AUTOMOTIVE HQ</div>
                      <div className="text-[10px] text-slate-500 leading-relaxed font-mono">
                        10, Avenue Edouard Serres - B.P. 60112<br />
                        31772 COLOMIERS Cedex, FRANCE<br />
                        Tél. +33 (0)5 61 17 61 61
                      </div>
                    </div>
                  </div>
                </div>

                {/* LEGEND BLOCK */}
                <div className="bg-slate-900 border border-slate-800 p-5 space-y-4 font-mono max-w-4xl mx-auto">
                  <div className="border-b border-slate-800 pb-2">
                    <span className="text-xs font-bold text-rose-400 uppercase tracking-widest flex items-center gap-1.5">
                      <Layers className="h-4 w-4" /> Critères d'Acceptation & Légendes Métrologiques
                    </span>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4 items-center">
                    <div className="space-y-3 text-[11px] text-slate-300 leading-relaxed">
                      <p>
                        Le rapport généré suit scrupuleusement la charte de certification de qualité d'<strong>ACTIA Automotive</strong>.
                      </p>
                      {mode === 'CAPABILITY' ? (
                        <div className="space-y-2 bg-slate-950 p-3 border border-slate-800">
                          <div className="font-bold text-xs text-rose-400 uppercase tracking-wider">CRITÈRES DE CAPABILITÉ (Cp/Cpk) :</div>
                          <ul className="list-disc pl-4 space-y-1">
                            <li><span className="text-white font-bold">Standard :</span> Le procédé est validé si le <span className="text-emerald-400 font-bold">Cpk ≥ {config.cpkLimNormal}</span>.</li>
                            <li><span className="text-white font-bold">Caractéristique Spéciale (CS / SC) :</span> Procédé critique exigeant un <span className="text-emerald-400 font-bold">Cpk ≥ {config.cpkLimSc}</span>.</li>
                            <li>Cp évalue la capabilité potentielle du procédé (centrage parfait).</li>
                            <li>Cpk évalue la capabilité réelle prenant en compte l'écart de centrage.</li>
                          </ul>
                        </div>
                      ) : (
                        <div className="space-y-2 bg-slate-950 p-3 border border-slate-800">
                          <div className="font-bold text-xs text-rose-400 uppercase tracking-wider">CRITÈRES DE RÉPÉTABILITÉ GRR (MSA) :</div>
                          <ul className="list-disc pl-4 space-y-1">
                            <li><span className="text-emerald-400 font-bold">%GRR &lt; 10% :</span> Système de mesure excellent, pleinement conforme.</li>
                            <li><span className="text-amber-400 font-bold">10% ≤ %GRR ≤ 30% :</span> Système acceptable selon l'importance.</li>
                            <li><span className="text-rose-400 font-bold">%GRR &gt; 30% :</span> Système inacceptable (à réviser/étalonner).</li>
                            <li><span className="text-white font-bold">NDC (Distinct Categories) :</span> Doit être <span className="text-emerald-400 font-bold">≥ {config.ndcLim}</span>.</li>
                          </ul>
                        </div>
                      )}
                    </div>

                    <div className="flex flex-col items-center justify-center p-4 bg-slate-950 border border-slate-850 space-y-4">
                      <div className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">Visualisation de la Légende de Calcul</div>
                      
                      {mode === 'CAPABILITY' ? (
                        <div className="relative group">
                          <img 
                            src="/actipa-assets/legendeCPK2.png" 
                            className="max-h-[160px] w-auto bg-white p-2 border border-slate-800 rounded shadow-md" 
                            referrerPolicy="no-referrer"
                            alt="Légende CPK" 
                          />
                        </div>
                      ) : (
                        <div className="grid grid-cols-2 gap-3">
                          <div className="flex flex-col items-center space-y-1">
                            <span className="text-[9px] text-slate-500 font-bold">LÉGENDE GRR</span>
                            <img 
                              src="/actipa-assets/legendeGRR2.png" 
                              className="max-h-[85px] w-auto bg-white p-1 border border-slate-800 rounded" 
                              referrerPolicy="no-referrer"
                              alt="Légende GRR" 
                            />
                          </div>
                          <div className="flex flex-col items-center space-y-1">
                            <span className="text-[9px] text-slate-500 font-bold">CRITÈRE NDC</span>
                            <img 
                              src="/actipa-assets/legendeNDC2.PNG" 
                              className="max-h-[85px] w-auto bg-white p-1 border border-slate-800 rounded" 
                              referrerPolicy="no-referrer"
                              alt="Légende NDC" 
                            />
                          </div>
                          <div className="col-span-2 flex flex-col items-center space-y-1 pt-1 border-t border-slate-800 w-full">
                            <span className="text-[9px] text-slate-500 font-bold">COEFFICIENTS CONSTANTES GRR (K2)</span>
                            <img 
                              src="/actipa-assets/legendeK2.png" 
                              className="max-h-[60px] w-auto bg-white p-1 border border-slate-800 rounded" 
                              referrerPolicy="no-referrer"
                              alt="Légende K2" 
                            />
                          </div>
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="flex justify-end pt-2 border-t border-slate-800/60">
                    <button
                      onClick={triggerExcelDownload}
                      className="bg-emerald-600 hover:bg-emerald-500 text-slate-950 font-black px-6 py-3 text-xs uppercase tracking-wider transition-all flex items-center gap-2 border border-emerald-500 cursor-pointer"
                    >
                      <FileSpreadsheet className="h-4 w-4" /> Télécharger ce Rapport Excel Certifié
                    </button>
                  </div>
                </div>
              </div>
            ) : (
              /* MAIN RESULTS TABLE */
              <div className="bg-slate-900 border border-slate-800 overflow-hidden shadow-xl">
                {/* Table controls */}
                <div className="p-5 border-b border-slate-950 flex flex-col sm:flex-row items-center justify-between gap-4 font-mono">
                  <div className="relative w-full sm:w-72">
                    <input
                      type="text"
                      placeholder="Filtrer par Test Point ID..."
                      value={searchQuery}
                      onChange={e => setSearchQuery(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-850 rounded-none p-2.5 pl-9 text-xs text-white placeholder-slate-600 focus:outline-none focus:border-rose-500"
                    />
                    <Search className="absolute left-3 top-3 h-4 w-4 text-slate-600" />
                  </div>

                  <button
                    onClick={triggerExcelDownload}
                    className="w-full sm:w-auto bg-emerald-600 hover:bg-emerald-500 text-slate-950 font-black px-6 py-2.5 rounded-none text-xs uppercase tracking-wider transition-all flex items-center justify-center gap-2 shadow-[0_0_15px_rgba(16,185,129,0.15)] border border-emerald-500 cursor-pointer"
                  >
                    <FileSpreadsheet className="h-4 w-4" /> Exporter Rapport Excel
                  </button>
                </div>

                {/* Scrollable table container */}
                <div className="overflow-x-auto">
                  {/* CAPABILITY TABLE */}
                  {mode === 'CAPABILITY' && (
                    <table className="w-full text-left font-mono text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-950 text-slate-400 uppercase text-[10px] tracking-widest border-b border-slate-850">
                          <th className="px-6 py-4 font-bold">Test Point ID</th>
                          <th className="px-6 py-4 font-bold">Spécifications (LSL / USL)</th>
                          <th className="px-6 py-4 font-bold text-right">Moyenne</th>
                          <th className="px-6 py-4 font-bold text-right">Écart-type (σ)</th>
                          <th className="px-6 py-4 font-bold text-right">Cp</th>
                          <th className="px-6 py-4 font-bold text-right">Cpk</th>
                          <th className="px-6 py-4 font-bold text-right">Centrage</th>
                          <th className="px-6 py-4 font-bold text-center">Statut</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-950 text-slate-300">
                        {filteredCapability.length === 0 ? (
                          <tr>
                            <td colSpan={8} className="px-6 py-8 text-center text-slate-500 italic">Aucun test point ne correspond à la recherche.</td>
                          </tr>
                        ) : (
                          filteredCapability.map((r, idx) => (
                            <tr key={idx} className="hover:bg-slate-950/40 transition-colors">
                              <td className="px-6 py-4 font-bold">
                                <div>{r.testPointId}</div>
                                {r.testStepName && <span className="text-[10px] text-slate-500">{r.testStepName}</span>}
                              </td>
                              <td className="px-6 py-4 text-slate-400">
                                {r.lsl} {r.unit} à {r.usl} {r.unit}
                              </td>
                              <td className="px-6 py-4 text-right">{r.mean.toFixed(4)}</td>
                              <td className="px-6 py-4 text-right text-slate-400">{r.sigma.toFixed(5)}</td>
                              <td className="px-6 py-4 text-right font-bold text-white">
                                {r.cp === 9999 ? "N/A" : r.cp.toFixed(3)}
                              </td>
                              <td className={`px-6 py-4 text-right font-black ${
                                r.status === 'OK' ? 'text-emerald-400' : 'text-rose-400'
                              }`}>
                                {r.cpk === 9999 ? "N/A" : r.cpk.toFixed(3)}
                              </td>
                              <td className="px-6 py-4 text-right text-slate-400">{r.centering.toFixed(3)}</td>
                              <td className="px-6 py-4 text-center">
                                <span className={`px-2.5 py-1 text-[10px] font-bold tracking-wider uppercase inline-block border ${
                                  r.status === 'OK' 
                                    ? 'bg-emerald-950 text-emerald-400 border-emerald-900' 
                                    : r.status === 'NOK'
                                    ? 'bg-rose-950 text-rose-400 border-rose-900'
                                    : 'bg-slate-950 text-slate-500 border-slate-850'
                                }`}>
                                  {r.status}
                                </span>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  )}

                  {/* REPEATABILITY GRR TABLE */}
                  {mode === 'REPEATABILITY' && (
                    <table className="w-full text-left font-mono text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-950 text-slate-400 uppercase text-[10px] tracking-widest border-b border-slate-850">
                          <th className="px-6 py-4 font-bold">Test Point ID</th>
                          <th className="px-6 py-4 font-bold text-right">Rbar (Portées)</th>
                          <th className="px-6 py-4 font-bold text-right">EV (Équipement)</th>
                          <th className="px-6 py-4 font-bold text-right">AV (Appareil)</th>
                          <th className="px-6 py-4 font-bold text-right">GRR Global</th>
                          <th className="px-6 py-4 font-bold text-right">%GRR</th>
                          <th className="px-6 py-4 font-bold text-right">NDC</th>
                          <th className="px-6 py-4 font-bold text-center">Statut</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-950 text-slate-300">
                        {filteredRepeatability.length === 0 ? (
                          <tr>
                            <td colSpan={8} className="px-6 py-8 text-center text-slate-500 italic">Aucun test point ne correspond à la recherche.</td>
                          </tr>
                        ) : (
                          filteredRepeatability.map((r, idx) => (
                            <tr key={idx} className="hover:bg-slate-950/40 transition-colors">
                              <td className="px-6 py-4 font-bold">
                                <div>{r.testPointId}</div>
                                {r.testStepName && <span className="text-[10px] text-slate-500">{r.testStepName}</span>}
                              </td>
                              <td className="px-6 py-4 text-right text-slate-400">{r.rBar.toFixed(5)}</td>
                              <td className="px-6 py-4 text-right text-slate-400">{r.ev.toFixed(5)}</td>
                              <td className="px-6 py-4 text-right text-slate-400">{r.av.toFixed(5)}</td>
                              <td className="px-6 py-4 text-right font-bold text-white">{r.grr.toFixed(5)}</td>
                              <td className={`px-6 py-4 text-right font-black ${
                                r.status === 'OK' ? 'text-emerald-400' : 'text-rose-400'
                              }`}>
                                {(r.percentGrr * 100).toFixed(2)}%
                              </td>
                              <td className="px-6 py-4 text-right font-bold text-slate-200">
                                {r.ndc === 9999 ? "N/A" : r.ndc.toFixed(2)}
                              </td>
                              <td className="px-6 py-4 text-center">
                                <span className={`px-2.5 py-1 text-[10px] font-bold tracking-wider uppercase inline-block border ${
                                  r.status === 'OK' 
                                    ? 'bg-emerald-950 text-emerald-400 border-emerald-900' 
                                    : r.status === 'NOK'
                                    ? 'bg-rose-950 text-rose-400 border-rose-900'
                                    : 'bg-slate-950 text-slate-500 border-slate-850'
                                }`}>
                                  {r.status}
                                </span>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  )}

                  {/* AIRBUS EV% TABLE */}
                  {mode === 'AIRBUS_EV' && (
                    <table className="w-full text-left font-mono text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-950 text-slate-400 uppercase text-[10px] tracking-widest border-b border-slate-850">
                          <th className="px-6 py-4 font-bold">Test Point ID</th>
                          <th className="px-6 py-4 font-bold text-right">Moyenne</th>
                          <th className="px-6 py-4 font-bold text-right">Écart-Type (σ)</th>
                          <th className="px-6 py-4 font-bold text-right">Range Max-Min</th>
                          <th className="px-6 py-4 font-bold text-right">EV Absolue</th>
                          <th className="px-6 py-4 font-bold text-right">EV % (Airbus)</th>
                          <th className="px-6 py-4 font-bold text-center">Statut</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-950 text-slate-300">
                        {filteredAirbus.length === 0 ? (
                          <tr>
                            <td colSpan={7} className="px-6 py-8 text-center text-slate-500 italic">Aucun test point ne correspond à la recherche.</td>
                          </tr>
                        ) : (
                          filteredAirbus.map((r, idx) => (
                            <tr key={idx} className="hover:bg-slate-950/40 transition-colors">
                              <td className="px-6 py-4 font-bold">
                                <div>{r.testPointId}</div>
                                {r.testStepName && <span className="text-[10px] text-slate-500">{r.testStepName}</span>}
                              </td>
                              <td className="px-6 py-4 text-right">{r.mean.toFixed(4)}</td>
                              <td className="px-6 py-4 text-right text-slate-400">{r.sigma.toFixed(5)}</td>
                              <td className="px-6 py-4 text-right text-slate-400">{r.range.toFixed(4)}</td>
                              <td className="px-6 py-4 text-right font-bold text-white">{r.evAbs.toFixed(5)}</td>
                              <td className={`px-6 py-4 text-right font-black ${
                                r.status === 'OK' ? 'text-emerald-400' : 'text-rose-400'
                              }`}>
                                {(r.evPercent * 100).toFixed(2)}%
                              </td>
                              <td className="px-6 py-4 text-center">
                                <span className={`px-2.5 py-1 text-[10px] font-bold tracking-wider uppercase inline-block border ${
                                  r.status === 'OK' 
                                    ? 'bg-emerald-950 text-emerald-400 border-emerald-900' 
                                    : r.status === 'NOK'
                                    ? 'bg-rose-950 text-rose-400 border-rose-900'
                                    : 'bg-slate-950 text-slate-500 border-slate-850'
                                }`}>
                                  {r.status}
                                </span>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      {/* NO FILES ENCOURAGING TEXT */}
      {!analysisRan && files.length === 0 && (
        <div className="bg-slate-900/10 border border-slate-900/60 p-8 flex flex-col items-center justify-center text-center font-mono py-12">
          <AlertCircle className="h-8 w-8 text-rose-500/80 mb-3" />
          <div className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-1">
            Aucune donnée à analyser
          </div>
          <p className="text-[11px] text-slate-600 max-w-md">
            Importez des fichiers de logs ci-dessus pour lancer les calculs métrologiques de capabilité (AT242) ou de répétabilité GRR (AT243).
          </p>
        </div>
      )}
    </div>
  );
}
