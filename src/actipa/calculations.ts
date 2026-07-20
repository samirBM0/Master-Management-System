import { 
  TestStepAggregated, 
  CapabilityResult, 
  RepeatabilityResult, 
  AirBusEVResult, 
  ActipaCalculationConfig 
} from "./types";

export function calculateCapability(
  step: TestStepAggregated,
  config: ActipaCalculationConfig,
  isSc: boolean = false
): CapabilityResult {
  const { lsl, usl, nominal, mean, sigma, values } = step;
  const n = values.length;

  if (n === 0) {
    return {
      testPointId: step.testPointId,
      testStepName: step.testStepName,
      unit: step.unit,
      lsl,
      usl,
      nominal,
      mean: 0,
      sigma: 0,
      cp: 0,
      cpk: 0,
      centering: 0,
      status: 'ND',
      valuesCount: 0
    };
  }

  const targetCpk = isSc ? config.cpkLimSc : config.cpkLimNormal;

  if (sigma === 0) {
    return {
      testPointId: step.testPointId,
      testStepName: step.testStepName,
      unit: step.unit,
      lsl,
      usl,
      nominal,
      mean,
      sigma,
      cp: 9999,
      cpk: 9999,
      centering: 0,
      status: 'OK',
      valuesCount: n
    };
  }

  const rangeTol = usl - lsl;
  let cp = 9999;
  if (rangeTol > 0) {
    cp = rangeTol / (6 * sigma);
  }

  const cpkLower = (mean - lsl) / (3 * sigma);
  const cpkUpper = (usl - mean) / (3 * sigma);
  let cpk = Math.min(cpkLower, cpkUpper);

  if (cpk > 1000) {
    cpk = 1000;
  } else if (cpk < -50) {
    cpk = -50;
  }

  if (cp > 1000) {
    cp = 1000;
  }

  let centering = 0;
  if (rangeTol > 0) {
    centering = (2 * (mean - nominal)) / rangeTol;
  }

  const status = cpk >= targetCpk ? 'OK' : 'NOK';

  return {
    testPointId: step.testPointId,
    testStepName: step.testStepName,
    unit: step.unit,
    lsl,
    usl,
    nominal,
    mean,
    sigma,
    cp,
    cpk,
    centering,
    status,
    valuesCount: n
  };
}

export interface GrrFileAssignment {
  filename: string;
  operatorIndex: number; // 0 or 1
  partIndex: number;     // 0 or 1
  trialIndex: number;    // 0, 1 or 2
}

export function assignGrrMetadata(filename: string, index: number, total: number): GrrFileAssignment {
  const lower = filename.toLowerCase();
  
  let operatorIndex = 0;
  if (lower.includes("op2") || lower.includes("oper2") || lower.includes("operateur2") || lower.includes("op_2") || lower.includes("operator2") || lower.includes("tech_b") || lower.includes("techb") || lower.includes("o2")) {
    operatorIndex = 1;
  } else if (lower.includes("op1") || lower.includes("oper1") || lower.includes("operateur1") || lower.includes("op_1") || lower.includes("operator1") || lower.includes("tech_a") || lower.includes("techa") || lower.includes("o1")) {
    operatorIndex = 0;
  } else {
    // Fallback: first half of files are Operator 1, second half are Operator 2
    operatorIndex = index < (total / 2) ? 0 : 1;
  }

  let partIndex = 0;
  if (lower.includes("p2") || lower.includes("part2") || lower.includes("piece2") || lower.includes("part_2") || lower.includes("piece_2") || lower.includes("partb") || lower.includes("part b")) {
    partIndex = 1;
  } else if (lower.includes("p1") || lower.includes("part1") || lower.includes("piece1") || lower.includes("part_1") || lower.includes("piece_1") || lower.includes("parta") || lower.includes("part a")) {
    partIndex = 0;
  } else {
    // Fallback: within operator's group, split into part 1 and part 2
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
    // Fallback: modulo based on 3 trials
    trialIndex = index % 3;
  }

  return {
    filename,
    operatorIndex,
    partIndex,
    trialIndex
  };
}

export function calculateRepeatability(
  step: TestStepAggregated,
  fileAssignments: GrrFileAssignment[],
  config: ActipaCalculationConfig,
  isSc: boolean = false
): RepeatabilityResult {
  const { lsl, usl, mean, sigma, values } = step;
  const n = values.length;

  if (n < 4) {
    return {
      testPointId: step.testPointId,
      testStepName: step.testStepName,
      unit: step.unit,
      lsl,
      usl,
      mean,
      rBar: 0,
      ev: 0,
      av: 0,
      grr: 0,
      percentGrr: 0,
      pv: 0,
      ndc: 0,
      status: 'ND',
      valuesCount: n
    };
  }

  // Group values by Operator and Part
  // matrix[operatorIndex][partIndex] = array of trial values
  const matrix: number[][][] = [
    [[], []], // Operator 0: Part 0, Part 1
    [[], []]  // Operator 1: Part 0, Part 1
  ];

  for (let i = 0; i < fileAssignments.length; i++) {
    const assign = fileAssignments[i];
    const val = values[i];
    if (val !== undefined && assign) {
      matrix[assign.operatorIndex][assign.partIndex].push(val);
    }
  }

  // Calculate ranges for each (operator, part)
  const ranges: number[] = [];
  const operatorMeans: number[][] = [[], []]; // operatorMeans[op][part] = mean
  const allOp1Values: number[] = [];
  const allOp2Values: number[] = [];

  for (let op = 0; op < 2; op++) {
    for (let part = 0; part < 2; part++) {
      const trials = matrix[op][part];
      if (trials.length > 0) {
        const max = Math.max(...trials);
        const min = Math.min(...trials);
        ranges.push(max - min);
        
        const sum = trials.reduce((s, v) => s + v, 0);
        operatorMeans[op].push(sum / trials.length);

        if (op === 0) {
          allOp1Values.push(...trials);
        } else {
          allOp2Values.push(...trials);
        }
      }
    }
  }

  const rBar = ranges.length > 0 ? ranges.reduce((s, v) => s + v, 0) / ranges.length : 0;
  const ev = rBar * config.k1_m3;

  // Xdiff = difference between operators' grand means
  const meanOp1 = allOp1Values.length > 0 ? allOp1Values.reduce((s, v) => s + v, 0) / allOp1Values.length : mean;
  const meanOp2 = allOp2Values.length > 0 ? allOp2Values.reduce((s, v) => s + v, 0) / allOp2Values.length : mean;
  const xDiff = Math.abs(meanOp1 - meanOp2);

  const avSq = Math.pow(xDiff * config.k2, 2) - (Math.pow(ev, 2) / 6);
  const av = avSq < 0 ? 0 : Math.sqrt(avSq);

  const grr = Math.sqrt(Math.pow(ev, 2) + Math.pow(av, 2));

  let rangeTol = usl - lsl;
  if (isNaN(rangeTol) || rangeTol <= 0) {
    rangeTol = 6 * sigma;
  }

  const tv = rangeTol > 1e50 ? 2 * Math.min(mean - lsl, usl - mean) : rangeTol / 6;
  const percentGrr = tv === 0 ? 0 : grr / tv;

  const pvSq = Math.pow(tv, 2) - Math.pow(grr, 2);
  const pv = pvSq < 0 ? 0 : Math.sqrt(pvSq);

  const ndc = grr === 0 ? 9999 : Math.sqrt(2) * (pv / grr);

  const targetMsa = isSc ? config.msaLimSc : config.msaLimNormal;
  const status = (percentGrr <= targetMsa && ndc >= config.ndcLim) ? 'OK' : 'NOK';

  return {
    testPointId: step.testPointId,
    testStepName: step.testStepName,
    unit: step.unit,
    lsl,
    usl,
    mean,
    rBar,
    ev,
    av,
    grr,
    percentGrr,
    pv,
    ndc,
    status,
    valuesCount: n
  };
}

export function calculateAirbusEV(
  step: TestStepAggregated,
  config: ActipaCalculationConfig,
  isSc: boolean = false
): AirBusEVResult {
  const { lsl, usl, mean, sigma, values } = step;
  const n = values.length;

  if (n < 2) {
    return {
      testPointId: step.testPointId,
      testStepName: step.testStepName,
      unit: step.unit,
      lsl,
      usl,
      mean,
      sigma,
      range: 0,
      evAbs: 0,
      evPercent: 0,
      status: 'ND',
      valuesCount: n
    };
  }

  const max = Math.max(...values);
  const min = Math.min(...values);
  const range = max - min;
  const evAbs = range * config.k1_airbus;

  let rangeTol = usl - lsl;
  if (isNaN(rangeTol) || rangeTol <= 0) {
    rangeTol = 6 * sigma;
  }

  const toleranceVariation = rangeTol > 1e50 
    ? Math.min(2 * (mean - lsl), 2 * (usl - mean)) 
    : rangeTol / 6;

  if (toleranceVariation === 0) {
    return {
      testPointId: step.testPointId,
      testStepName: step.testStepName,
      unit: step.unit,
      lsl,
      usl,
      mean,
      sigma,
      range,
      evAbs,
      evPercent: 0,
      status: 'ND',
      valuesCount: n
    };
  }

  const evPercentSigma = (sigma * config.k1_airbus) / toleranceVariation;
  const evPercentRange = evAbs / toleranceVariation;
  const evPercent = Math.min(evPercentSigma, evPercentRange, 1);

  const targetEv = isSc ? config.evLimSc : config.evLimNormal;
  const status = evPercent <= targetEv ? 'OK' : 'NOK';

  return {
    testPointId: step.testPointId,
    testStepName: step.testStepName,
    unit: step.unit,
    lsl,
    usl,
    mean,
    sigma,
    range,
    evAbs,
    evPercent,
    status,
    valuesCount: n
  };
}
