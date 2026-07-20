export interface ActiaLogMeasurement {
  testPointId: string;
  unit: string;
  status: string;
  value: number;
  lsl?: number;
  usl?: number;
  category?: string;
  testStepName?: string;
}

export interface ParsedLogFile {
  filename: string;
  productRef: string;
  status: 'PASS' | 'FAIL';
  measurements: ActiaLogMeasurement[];
  testName?: string;
}

export interface TestStepAggregated {
  testPointId: string;
  testStepName: string;
  unit: string;
  lsl: number;
  usl: number;
  nominal: number;
  values: number[];
  mean: number;
  sigma: number;
}

export interface CapabilityResult {
  testPointId: string;
  testStepName: string;
  unit: string;
  lsl: number;
  usl: number;
  nominal: number;
  mean: number;
  sigma: number;
  cp: number;
  cpk: number;
  centering: number;
  status: 'OK' | 'NOK' | 'ND';
  valuesCount: number;
}

export interface RepeatabilityResult {
  testPointId: string;
  testStepName: string;
  unit: string;
  lsl: number;
  usl: number;
  mean: number;
  rBar: number;
  ev: number;
  av: number;
  grr: number;
  percentGrr: number;
  pv: number;
  ndc: number;
  status: 'OK' | 'NOK' | 'ND';
  valuesCount: number;
}

export interface AirBusEVResult {
  testPointId: string;
  testStepName: string;
  unit: string;
  lsl: number;
  usl: number;
  mean: number;
  sigma: number;
  range: number;
  evAbs: number;
  evPercent: number;
  status: 'OK' | 'NOK' | 'ND';
  valuesCount: number;
}

export interface ActipaCalculationConfig {
  cpkLimNormal: number;
  cpkLimSc: number;
  msaLimNormal: number;
  msaLimSc: number;
  ndcLim: number;
  evLimNormal: number;
  evLimSc: number;
  k1_m3: number;
  k2: number;
  k1_airbus: number;
}

export const DEFAULT_ACTIA_CONFIG: ActipaCalculationConfig = {
  cpkLimNormal: 1.33,
  cpkLimSc: 1.67,
  msaLimNormal: 0.30,
  msaLimSc: 0.10,
  ndcLim: 5,
  evLimNormal: 0.30,
  evLimSc: 0.10,
  k1_m3: 0.5908,
  k2: 0.7071,
  k1_airbus: 0.8623
};
