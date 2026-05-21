export function achClass(pct: number): string {
  if (!pct || pct === 0) return 'ach-grey';
  if (pct > 120) return 'ach-dark-red';
  if (pct > 100) return 'ach-red';
  if (pct >= 80) return 'ach-orange';
  if (pct >= 50) return 'ach-light-green';
  return 'ach-dark-green';
}

export function fmt(v: number | undefined | null, d = 2): string {
  if (v == null || isNaN(v)) return '0';
  return v.toLocaleString('en-IN', { minimumFractionDigits: d, maximumFractionDigits: d });
}

export function fmtPct(v: number | undefined | null): string {
  if (v == null || isNaN(v)) return '0.0%';
  return v.toFixed(1) + '%';
}

export interface Row {
  month: string; year: string | number; lastModDate: string; lastModTime: string;
  itemCode: string; bomType: string; itemGroup: string; itemName: string;
  itemParent: string; convFactor: number; projUnits: number; customer: string;
  projKg: number; custGroup: string; origin: string; itemType: string; newMIS: string;
  dailyKg: number; dailyUnits: number; expKg: number; expUnits: number;
  soKg: number; soUnits: number; diffKg: number; diffUnits: number;
  achKg: number; achUnits: number; runRateKg: number; runRateUnits: number;
  soVsProj: number; soLeftPct: number; soPctPerDay: number; daysToCover: number;
  isProj: boolean; isUnproj: boolean;
}

export interface Summary {
  name: string; projKg: number; projUnits: number; expKg: number; expUnits: number;
  soKg: number; soUnits: number; achKg: number; achUnits: number;
  diffKg: number; diffUnits: number; count: number;
}

export interface CleaningRuleStat {
  rule: string;
  removedRows: number;
  removedKg: number;
}

export interface DashData {
  config: {
    daysInMonth: number; dateFrom: string; dateTo: string; daysElapsed: number;
    soRowsRaw: number; soRowsClean: number; totalRows: number;
    projFile: string; soFile: string;
  };
  filters: {
    custGroups: string[]; origins: string[]; itemGroups: string[];
    customers: string[]; newMIS: string[];
  };
  cleaningStats: CleaningRuleStat[];
  rows: Row[];
  summaryIG: Summary[]; summaryOrigin: Summary[];
  summaryCG: Summary[]; summaryMIS: Summary[];
  top20Kg: Row[]; top20Units: Row[]; bot20Kg: Row[]; bot20Units: Row[];
}
