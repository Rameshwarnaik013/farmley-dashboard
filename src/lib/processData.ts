import * as XLSX from 'xlsx';
import { DashData, Row, Summary, CleaningRuleStat } from './helpers';

const VALID_IT = ['ashish.k@farmley.com', 'abhishek.ku@farmley.com'];

function round2(v: number) { return Math.round(v * 100) / 100; }

function parseDate(v: unknown): Date {
  if (v instanceof Date) return v;
  if (typeof v === 'number') return new Date((v - 25569) * 86400000);
  return new Date(v as string);
}

// Trim and normalize string fields
function str(v: unknown): string {
  if (v == null) return '';
  return String(v).trim();
}

// Build a normalized join key
function joinKey(cg: string, cust: string, origin: string, ic: string): string {
  return `${cg}|${cust}|${origin}|${ic}`;
}

interface RawRow { [key: string]: unknown; }

interface CleanResult { cleaned: RawRow[]; stats: CleaningRuleStat[]; }

function cleanSO(rows: RawRow[]): CleanResult {
  const rules: { name: string; test: (r: RawRow) => boolean }[] = [
    { name: 'Created By Administrator', test: r => str(r['Sales Order Created By']) === 'Administrator' },
    { name: 'WF State: Internal Transfer (invalid creator)', test: r => str(r['Workflow State']) === 'Internal Transfer' && !VALID_IT.includes(str(r['Sales Order Created By'])) },
    { name: 'WF State: On Hold / Rejected / Cancelled', test: r => ['On Hold', 'Rejected', 'Cancelled'].includes(str(r['Workflow State'])) },
    { name: 'Purpose: Documentation', test: r => str(r['Purpose']) === 'Documentation' },
    { name: 'SO Status: Cancelled / On Hold / Rejected', test: r => ['Cancelled', 'On Hold', 'Rejected'].includes(str(r['Sales Order Status'])) },
    { name: 'SO Status: Internal Transfer (invalid creator)', test: r => str(r['Sales Order Status']) === 'Internal Transfer' && !VALID_IT.includes(str(r['Sales Order Created By'])) },
    { name: 'Customer starts with Sample Order', test: r => str(r['Customer']).startsWith('Sample Order') },
    { name: 'Item Type: Bulk', test: r => str(r['Item Type']) === 'Bulk' },
    { name: 'Returned Qty > 1', test: r => (Number(r['Returned Qty']) || 0) > 1 },
  ];

  const stats: CleaningRuleStat[] = rules.map(r => ({ rule: r.name, removedRows: 0, removedKg: 0 }));

  const cleaned = rows.filter(r => {
    const kg = Number(r['Stock Qty']) || 0;
    for (let i = 0; i < rules.length; i++) {
      if (rules[i].test(r)) {
        stats[i].removedRows++;
        stats[i].removedKg += kg;
        return false;
      }
    }
    return true;
  });

  // Round stats
  stats.forEach(s => { s.removedKg = Math.round(s.removedKg * 100) / 100; });

  return { cleaned, stats: stats.filter(s => s.removedRows > 0) };
}

export function aggGroup(rows: Row[], groupKey: keyof Row): Summary[] {
  const map: Record<string, Summary> = {};
  rows.forEach(r => {
    const g = (r[groupKey] as string) || 'Unknown';
    if (!map[g]) map[g] = { name: g, projKg: 0, projUnits: 0, expKg: 0, expUnits: 0, soKg: 0, soUnits: 0, achKg: 0, achUnits: 0, diffKg: 0, diffUnits: 0, count: 0 };
    map[g].projKg += r.projKg;
    map[g].projUnits += r.projUnits;
    map[g].expKg += r.expKg;
    map[g].expUnits += r.expUnits;
    map[g].soKg += r.soKg;
    map[g].soUnits += r.soUnits;
    map[g].count++;
  });
  return Object.values(map).map(g => ({
    ...g,
    projKg: round2(g.projKg), projUnits: round2(g.projUnits),
    expKg: round2(g.expKg), expUnits: round2(g.expUnits),
    soKg: round2(g.soKg), soUnits: round2(g.soUnits),
    achKg: round2(g.expKg > 0 ? (g.soKg / g.expKg) * 100 : 0),
    achUnits: round2(g.expUnits > 0 ? (g.soUnits / g.expUnits) * 100 : 0),
    diffKg: round2(g.expKg - g.soKg),
    diffUnits: round2(g.expUnits - g.soUnits),
  })).sort((a, b) => b.achKg - a.achKg);
}

function unique(arr: Row[], key: keyof Row): string[] {
  return [...new Set(arr.map(r => r[key] as string).filter(Boolean))].sort();
}

export function processExcelFiles(projBuffer: ArrayBuffer, soBuffer: ArrayBuffer, daysInMonth: number): DashData {
  // Read Excel files
  const projWb = XLSX.read(projBuffer, { type: 'array' });
  const soWb = XLSX.read(soBuffer, { type: 'array' });

  const projRaw: RawRow[] = XLSX.utils.sheet_to_json(projWb.Sheets[projWb.SheetNames[0]]);
  let soRaw: RawRow[] = XLSX.utils.sheet_to_json(soWb.Sheets[soWb.SheetNames[0]]);
  const soRawCount = soRaw.length;

  // Clean SO
  const cleanResult = cleanSO(soRaw);
  soRaw = cleanResult.cleaned;
  const cleaningStats = cleanResult.stats;

  // ─── Customer Group / Customer Remapping ───
  const INSTAMART_CUSTOMERS = new Set([
    'Moksh Enterprises Private Limited',
    'PJTJ Technologies Private Limited',
    'Cloudkart Ventures Private Limited',
    'Jupiter Kart Private Limited',
    'Cloudstore Retail Private Limited',
  ]);

  function remapCustomer(r: RawRow) {
    const cg = str(r['Customer Group']);
    const cust = str(r['Customer']);
    // Category A + Jhabak Marketing → Modern Trade
    if (cg === 'Category A' && cust === 'Jhabak Marketing') {
      r['Customer Group'] = 'Modern Trade';
    } else if (cg === 'Category A') {
      r['Customer'] = 'GT Retail';
    }
    // CPC KPKB → Customer = CPC KPKB
    if (cg === 'CPC KPKB') {
      r['Customer'] = 'CPC KPKB';
    }
    // Quick Commerce + specific customers → Instamart
    if (cg === 'Quick Commerce' && INSTAMART_CUSTOMERS.has(cust)) {
      r['Customer'] = 'Instamart';
    }
  }

  projRaw.forEach(remapCustomer);
  soRaw.forEach(remapCustomer);

  // ─── Date range — always from 1st of month to max SO date ───
  const soDates = soRaw.map(r => parseDate(r['Sales Order Date'])).filter(d => !isNaN(d.getTime()));
  soDates.sort((a, b) => a.getTime() - b.getTime());
  const dateTo = soDates[soDates.length - 1];
  const dateToDay = dateTo.getUTCDate() || dateTo.getDate();
  const dateToMonth = dateTo.getUTCMonth();
  const dateToYear = dateTo.getUTCFullYear() || dateTo.getFullYear();
  const dateFrom = new Date(Date.UTC(dateToYear, dateToMonth, 1));
  const daysElapsed = dateToDay;
  const dateFromStr = dateFrom.toISOString().slice(0, 10);
  const dateToStr = dateTo.toISOString().slice(0, 10);

  // ─── Build SO lookup: group by (CustGroup, Customer, Origin, ItemCode) → SUM(Stock Qty) ───
  const soMap: Record<string, { soKg: number; soUnits: number; newMIS: string; itemName: string; itemGroup: string; itemParent: string; convFactor: number; itemType: string }> = {};
  soRaw.forEach(r => {
    const cg = str(r['Customer Group']);
    const cust = str(r['Customer']);
    const origin = str(r['Origin']);
    const ic = str(r['Item Code']);
    const key = joinKey(cg, cust, origin, ic);

    if (!soMap[key]) {
      soMap[key] = {
        soKg: 0,
        soUnits: 0,
        newMIS: str(r['NEW MIS ITEM GROUP']),
        itemName: str(r['Item Name']),
        itemGroup: str(r['Item Group']),
        itemParent: str(r['Parent Item']),
        convFactor: Number(r['Conversion Factor']) || 0,
        itemType: str(r['Item Type']),
      };
    }
    // SUM up Stock Qty (KGs) and Qty (Units) for each group
    soMap[key].soKg += (Number(r['Stock Qty']) || 0);
    soMap[key].soUnits += (Number(r['Qty']) || 0);
    if (!soMap[key].newMIS && r['NEW MIS ITEM GROUP']) soMap[key].newMIS = str(r['NEW MIS ITEM GROUP']);
  });

  // SO lookup by ItemCode for New MIS fallback
  const soItemMIS: Record<string, string> = {};
  soRaw.forEach(r => {
    const ic = str(r['Item Code']);
    if (r['NEW MIS ITEM GROUP'] && !soItemMIS[ic]) soItemMIS[ic] = str(r['NEW MIS ITEM GROUP']);
  });

  // ─── Build enriched rows via OUTER JOIN on (CustGroup, Customer, Origin, ItemCode) ───
  const rows: Row[] = [];
  const projKeys = new Set<string>();

  projRaw.forEach(r => {
    const cg = str(r['Customer Group']);
    const cust = str(r['Customer']);
    const origin = str(r['Origin']);
    const ic = str(r['Item Code']);
    const key = joinKey(cg, cust, origin, ic);
    projKeys.add(key);

    const projKg = Number(r['Total KGs']) || 0;
    const projUnits = Number(r['Projection Units']) || 0;
    const dailyKg = projKg / daysInMonth;
    const dailyUnits = projUnits / daysInMonth;
    const expKg = dailyKg * daysElapsed;
    const expUnits = dailyUnits * daysElapsed;

    // Lookup SO aggregated value for this exact (CG, Customer, Origin, ItemCode) combination
    const so = soMap[key] || { soKg: 0, soUnits: 0, newMIS: '' };
    const achKg = expKg > 0 ? (so.soKg / expKg) * 100 : 0;
    const achUnits = expUnits > 0 ? (so.soUnits / expUnits) * 100 : 0;
    const soVsProj = projKg > 0 ? (so.soKg / projKg) * 100 : 0;
    const soLeftPct = 100 - soVsProj;
    const soPctPerDay = daysElapsed > 0 ? soVsProj / daysElapsed : 0;
    const daysToCover = soPctPerDay > 0 ? soLeftPct / soPctPerDay : 0;

    rows.push({
      month: str(r['Month']),
      year: str(r['Year']),
      lastModDate: str(r['Last Modified Date']),
      lastModTime: str(r['Last Modified Time']),
      itemCode: ic,
      bomType: str(r['BOM Item Type']),
      itemGroup: str(r['Item Group']),
      itemName: str(r['Item Name']),
      itemParent: str(r['Item Parent']),
      convFactor: Number(r['Conversion Factor']) || 0,
      projUnits: round2(projUnits),
      customer: cust,
      projKg: round2(projKg),
      custGroup: cg,
      origin: origin,
      itemType: str(r['Item Type']),
      newMIS: so.newMIS || soItemMIS[ic] || '',
      dailyKg: round2(dailyKg),
      dailyUnits: round2(dailyUnits),
      expKg: round2(expKg),
      expUnits: round2(expUnits),
      soKg: round2(so.soKg),
      soUnits: round2(so.soUnits),
      diffKg: round2(expKg - so.soKg),
      diffUnits: round2(expUnits - so.soUnits),
      achKg: round2(achKg),
      achUnits: round2(achUnits),
      runRateKg: round2(daysElapsed > 0 ? (so.soKg / daysElapsed) * daysInMonth : 0),
      runRateUnits: round2(daysElapsed > 0 ? (so.soUnits / daysElapsed) * daysInMonth : 0),
      soVsProj: round2(soVsProj),
      soLeftPct: round2(soLeftPct),
      soPctPerDay: round2(soPctPerDay),
      daysToCover: round2(daysToCover),
      isProj: true,
      isUnproj: false,
    });
  });

  // Add unprojected SO items (SO entries that had no matching projection)
  Object.entries(soMap).forEach(([key, so]) => {
    if (projKeys.has(key)) return;
    const [cg, cust, origin, ic] = key.split('|');
    rows.push({
      month: '', year: '', lastModDate: '', lastModTime: '',
      itemCode: ic,
      bomType: '',
      itemGroup: so.itemGroup,
      itemName: so.itemName,
      itemParent: so.itemParent,
      convFactor: so.convFactor,
      projUnits: 0, customer: cust, projKg: 0,
      custGroup: cg, origin: origin,
      itemType: so.itemType,
      newMIS: so.newMIS || '',
      dailyKg: 0, dailyUnits: 0, expKg: 0, expUnits: 0,
      soKg: round2(so.soKg), soUnits: round2(so.soUnits),
      diffKg: round2(-so.soKg), diffUnits: round2(-so.soUnits),
      achKg: 0, achUnits: 0,
      runRateKg: round2(daysElapsed > 0 ? (so.soKg / daysElapsed) * daysInMonth : 0),
      runRateUnits: round2(daysElapsed > 0 ? (so.soUnits / daysElapsed) * daysInMonth : 0),
      soVsProj: 0, soLeftPct: 0, soPctPerDay: 0, daysToCover: 0,
      isProj: false, isUnproj: true,
    });
  });

  // Pre-aggregate summaries (exclude Inter Company)
  const nonIC = rows.filter(r => r.custGroup !== 'Inter Company');
  const summaryIG = aggGroup(nonIC, 'itemGroup');
  const summaryOrigin = aggGroup(nonIC, 'origin');
  const summaryCG = aggGroup(nonIC, 'custGroup');
  const summaryMIS = aggGroup(nonIC, 'newMIS');

  // Leaderboards
  const withProj = nonIC.filter(r => r.isProj && r.expKg > 0);
  const top20Kg = [...withProj].sort((a, b) => b.achKg - a.achKg).filter(r => r.achKg > 100).slice(0, 20);
  const top20Units = [...withProj].sort((a, b) => b.achUnits - a.achUnits).filter(r => r.achUnits > 100).slice(0, 20);
  const bot20Kg = [...withProj].sort((a, b) => a.achKg - b.achKg).filter(r => r.achKg > 0 && r.achKg <= 100).slice(0, 20);
  const bot20Units = [...withProj].sort((a, b) => a.achUnits - b.achUnits).filter(r => r.achUnits > 0 && r.achUnits <= 100).slice(0, 20);

  return {
    config: {
      daysInMonth,
      dateFrom: dateFromStr,
      dateTo: dateToStr,
      daysElapsed,
      soRowsRaw: soRawCount,
      soRowsClean: soRaw.length,
      totalRows: rows.length,
      projFile: 'Uploaded Projection File',
      soFile: 'Uploaded SO File',
    },
    filters: {
      custGroups: unique(rows, 'custGroup'),
      origins: unique(rows, 'origin'),
      itemGroups: unique(rows, 'itemGroup'),
      customers: unique(rows, 'customer'),
      newMIS: unique(rows, 'newMIS'),
    },
    cleaningStats,
    rows,
    summaryIG, summaryOrigin, summaryCG, summaryMIS,
    top20Kg, top20Units, bot20Kg, bot20Units,
  };
}
