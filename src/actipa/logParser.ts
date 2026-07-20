import { ActiaLogMeasurement, ParsedLogFile, TestStepAggregated } from "./types";

export function cleanAccents(str: string): string {
  if (!str) return "";
  let s = str;
  // Fix specific patterns with replacement character (U+FFFD) or missing accent letters
  s = s.replace(/D\uFFFDcharge/gi, "Décharge");
  s = s.replace(/Dcharge/gi, "Décharge");
  s = s.replace(/Continuit\uFFFD/gi, "Continuité");
  s = s.replace(/Continuit/gi, "Continuité");
  s = s.replace(/Pr\uFFFDpar/gi, "Prépar");
  s = s.replace(/Prpar/gi, "Prépar");
  s = s.replace(/Temp\uFFFDr/gi, "Tempér");
  s = s.replace(/Tempr/gi, "Tempér");
  s = s.replace(/Mesur\uFFFD/gi, "Mesuré");
  s = s.replace(/Mesur$/gi, "Mesuré");
  s = s.replace(/M\uFFFDm/gi, "Mém");
  s = s.replace(/Mmoire/gi, "Mémoire");
  s = s.replace(/Entr\uFFFD/gi, "Entrée");
  s = s.replace(/Entre$/gi, "Entrée");
  s = s.replace(/R\uFFFDsultat/gi, "Résultat");
  s = s.replace(/Rsultat/gi, "Résultat");
  s = s.replace(/G\uFFFDn\uFFFDrateur/gi, "Générateur");
  s = s.replace(/Gnrateur/gi, "Générateur");
  s = s.replace(/R\uFFFDcepteur/gi, "Récepteur");
  s = s.replace(/Rcepteur/gi, "Récepteur");
  s = s.replace(/Pr\uFFFDcision/gi, "Précision");
  s = s.replace(/Prcision/gi, "Précision");
  s = s.replace(/Fr\uFFFDquence/gi, "Fréquence");
  s = s.replace(/Frquence/gi, "Fréquence");
  s = s.replace(/Intensit\uFFFD/gi, "Intensité");
  s = s.replace(/Intensit/gi, "Intensité");
  s = s.replace(/D\uFFFDfaut/gi, "Défaut");
  s = s.replace(/Dfaut/gi, "Défaut");
  s = s.replace(/V\uFFFDrification/gi, "Vérification");
  s = s.replace(/Vrification/gi, "Vérification");
  
  // Replace remaining \uFFFD characters with 'é'
  s = s.replace(/\uFFFD/g, "é");
  return s;
}

export function parseActiaLog(content: string, filename: string): ParsedLogFile {
  const lines = content.split(/\r?\n/);
  const measurements: ActiaLogMeasurement[] = [];
  let productRef = "";
  let overallStatus: 'PASS' | 'FAIL' = 'PASS';
  let testName = "";

  // Check for global PASS/FAIL indicators anywhere in the file
  const upperContent = content.toUpperCase();
  if (upperContent.includes("* FAIL *") || upperContent.includes("|M|") || upperContent.includes("|FAIL|") || upperContent.includes("|KO|")) {
    overallStatus = 'FAIL';
  } else if (upperContent.includes("* PASS *") || upperContent.includes("|B|") || upperContent.includes("|00|") || upperContent.includes("|OK|") || upperContent.includes("|PASS|")) {
    overallStatus = 'PASS';
  }

  for (const line of lines) {
    if (!line.trim()) continue;

    const rawParts = line.split('|');
    const parts = rawParts.map(p => p.trim());
    while (parts.length > 0 && parts[parts.length - 1] === "") {
      parts.pop();
    }

    const header = parts[0]?.toUpperCase();

    if (header === 'L24') {
      if (parts[1]) {
        productRef = parts[1];
      }
      if (parts[2]) {
        const stat = parts[2].toUpperCase();
        if (stat === 'M' || stat === 'FAIL' || stat === 'KO' || stat === 'F') {
          overallStatus = 'FAIL';
        } else if (stat === 'B' || stat === 'PASS' || stat === 'OK' || stat === 'P' || stat === '00') {
          overallStatus = 'PASS';
        }
      }
      if (parts.length >= 4) {
        testName = cleanAccents(parts[parts.length - 1] || "");
      }
    } else if (header === 'L301') {
      if (parts.length >= 5) {
        const testPointId = parts[1] || "";
        const unit = parts[2] || "";
        const status = parts[3] || "";
        const valStr = parts[4]?.replace(',', '.');
        const value = parseFloat(valStr);

        if (!isNaN(value)) {
          const lslStr = parts[5]?.replace(',', '.');
          const uslStr = parts[6]?.replace(',', '.');
          const lsl = lslStr ? parseFloat(lslStr) : undefined;
          const usl = uslStr ? parseFloat(uslStr) : undefined;

          const category = cleanAccents(parts[parts.length - 2] || "");
          const testStepName = cleanAccents(parts[parts.length - 1] || "");

          measurements.push({
            testPointId,
            unit,
            status,
            value,
            lsl: isNaN(lsl as number) ? undefined : lsl,
            usl: isNaN(usl as number) ? undefined : usl,
            category,
            testStepName
          });
        }
      }
    } else if (header === 'L302') {
      if (parts.length >= 5) {
        const testPointId = parts[1] || "";
        const unit = parts[2] || "";
        const status = parts[3] || "";
        const valStr = parts[4]?.replace(',', '.');
        const value = parseFloat(valStr);

        if (!isNaN(value)) {
          const category = cleanAccents(parts[parts.length - 2] || "");
          const testStepName = cleanAccents(parts[parts.length - 1] || "");

          measurements.push({
            testPointId,
            unit,
            status,
            value,
            category,
            testStepName
          });
        }
      }
    }
  }

  // If no L24 productRef found, try to extract from filename or use a default
  if (!productRef) {
    const match = filename.match(/^([A-Za-z0-9_-]+)/);
    productRef = match ? match[1] : "PRODUIT";
  }

  return {
    filename,
    productRef,
    status: overallStatus,
    measurements,
    testName
  };
}

export function aggregateTestSteps(parsedFiles: ParsedLogFile[]): TestStepAggregated[] {
  const groups: { [key: string]: {
    testPointId: string;
    testStepName: string;
    unit: string;
    lsls: number[];
    usls: number[];
    values: number[];
  }} = {};

  for (const file of parsedFiles) {
    for (const m of file.measurements) {
      const key = `${m.testPointId}::${m.testStepName || ""}`;
      if (!groups[key]) {
        groups[key] = {
          testPointId: m.testPointId,
          testStepName: m.testStepName || "",
          unit: m.unit,
          lsls: [],
          usls: [],
          values: []
        };
      }
      groups[key].values.push(m.value);
      if (m.lsl !== undefined && !isNaN(m.lsl)) {
        groups[key].lsls.push(m.lsl);
      }
      if (m.usl !== undefined && !isNaN(m.usl)) {
        groups[key].usls.push(m.usl);
      }
    }
  }

  const results: TestStepAggregated[] = [];

  for (const key in groups) {
    const g = groups[key];
    if (g.values.length === 0) continue;

    // Determine LSL & USL (most common value or average or first)
    const lsl = g.lsls.length > 0 ? g.lsls[0] : 0;
    const usl = g.usls.length > 0 ? g.usls[0] : 0;
    const nominal = (usl + lsl) / 2;

    const n = g.values.length;
    const mean = g.values.reduce((sum, v) => sum + v, 0) / n;

    let sigma = 0;
    if (n > 1) {
      const variance = g.values.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / (n - 1);
      sigma = Math.sqrt(variance);
    }

    results.push({
      testPointId: g.testPointId,
      testStepName: g.testStepName,
      unit: g.unit,
      lsl,
      usl,
      nominal,
      values: g.values,
      mean,
      sigma
    });
  }

  return results;
}
