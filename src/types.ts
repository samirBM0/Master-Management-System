export interface UserSession {
  role: 'admin' | 'technician';
  matricule: string;
  name: string;
  token: string;
}

export interface MasterItem {
  id: string;
  testeur: string;
  idMaster: string;
  refProduitMaster: string;
  numSerieProduitMaster: string;
  refCarteMaster: string;
  numSerieCarteMaster: string;
  dateCreation: string;
  commentaire1: string;
  statut: string; // e.g., 'Active', 'Obsolète', 'Endommagée'
  commentaire2: string;
  verif: string;   // e.g., 'OK', 'KO', 'FIN DE VIE'
}

export interface CapabilityMeasurement {
  trial: number;
  value: number;
}

export interface CapabilityCalculation {
  id: string;
  title: string;
  date: string;
  testeur: string;
  masterId: string;
  characteristic: string;
  nominal: number;
  lsl: number; // Lower Specification Limit (LSL)
  usl: number; // Upper Specification Limit (USL)
  measurements: number[];
  operator: string;
}

export interface HistoryLog {
  id: string;
  timestamp: string;
  user: string;
  role: 'admin' | 'technician';
  action: string;
  details: string;
}

