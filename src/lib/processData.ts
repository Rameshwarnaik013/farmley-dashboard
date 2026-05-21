import * as XLSX from 'xlsx';
import { DashData, Row, Summary } from './helpers';

const VALID_IT = ['ashish.k@farmley.com', 'abhishek.ku@farmley.com'];

function round2(v: number) { return Math.round(v * 100) / 100; }

function parseDate(v: unknown): Date {
  if (v instanceof Date) return v;
  if (typeof v === 'number') return new Date((v - 25569) * 86400000);
  return new Date(v as string);
}

interface RawRow { [key: string]: unknown; }

function cleanSO(rows: RawRow[]): RawRow[] {
  return rows.filter(r => {
    if (r['Sales Order Created By'] === 'Administrator') return false;
    if (r['Workflow State'] === 'Internal Transfer' && !VALID_IT.includes(r['Sales Order Created By'] as string)) return false;
    if (['On Hold', 'Rejected', 'Cancelled'].includes(r['Workflow State'] as string)) return false;
    if (r['Purpose'] === 'Documentation') return false;
    if (['Cancelled', 'On Hold', 'Rejected'].includes(r['Sales Order Status'] as string)) return false;
    if (r['Sales Order Status'] === 'Internal Transfer' && !VALID_IT.includes(r['Sales Order Created By'] as string)) return false;
    if (typeof r['Customer'] === 'string' && r['Customer'].startsWith('Sample Order')) return false;
    if (r['Item Type'] === 'Bulk') return false;
    if ((Number(r['Returned Qty']) || 0) > 1) return false;
    return true;
  });
}

function aggGroup(rows: Row[], groupKey: keyof Row): Summary[] {
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
  soRaw = cleanSO(soRaw);

  // Remap customer groups
  projRaw.forEach(r => {
    if (r['Customer Group'] === 'Category A' && r['Customer'] === 'Jhabak Marketing')
      r['Customer Group'] = 'Modern Trade';
  });
  soRaw.forEach(r => {
    if (r['Customer Group'] === 'Category A') r['Customer'] = 'GT Retail';
  });

  // Date range — always from 1st of month to max SO date
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

  // Build SO lookup
  const soMap: Record<string, { soKg: number; soUnits: number; newMIS: string }> = {};
  soRaw.forEach(r => {
    const key = `${r['Customer Group']}|${r['Customer']}|${r['Origin']}|${r['Item Code']}`;
    if (!soMap[key]) soMap[key] = { soKg: 0, soUnits: 0, newMIS: (r['NEW MIS ITEM GROUP'] as string) || '' };
    soMap[key].soKg += (Number(r['Stock Qty']) || 0);
    soMap[key].soUnits += (Number(r['Qty']) || 0);
    if (!soMap[key].newMIS && r['NEW MIS ITEM GROUP']) soMap[key].newMIS = r['NEW MIS ITEM GROUP'] as string;
  });

  // SO lookup by ItemCode for New MIS
  const soItemMIS: Record<string, string> = {};
  soRaw.forEach(r => {
    const ic = r['Item Code'] as string;
    if (r['NEW MIS ITEM GROUP'] && !soItemMIS[ic]) soItemMIS[ic] = r['NEW MIS ITEM GROUP'] as string;
  });

  // Build enriched projection rows
  const rows: Row[] = [];
  const projKeys = new Set<string>();

  projRaw.forEach(r => {
    const cg = (r['Customer Group'] as string) || '';
    const cust = (r['Customer'] as string) || '';
    const origin = (r['Origin'] as string) || '';
    const ic = (r['Item Code'] as string) || '';
    const key = `${cg}|${cust}|${origin}|${ic}`;
    projKeys.add(key);

    const projKg = Number(r['Total KGs']) || 0;
    const projUnits = Number(r['Projection Units']) || 0;
    const dailyKg = projKg / daysInMonth;
    const dailyUnits = projUnits / daysInMonth;
    const expKg = dailyKg * daysElapsed;
    const expUnits = dailyUnits * daysElapsed;

    const so = soMap[key] || { soKg: 0, soUnits: 0, newMIS: '' };
    const achKg = expKg > 0 ? (so.soKg / expKg) * 100 : 0;
    const achUnits = expUnits > 0 ? (so.soUnits / expUnits) * 100 : 0;

    rows.push({
      month: (r['Month'] as string) || '',
      year: (r['Year'] as string) || '',
      lastModDate: (r['Last Modified Date'] as string) || '',
      lastModTime: (r['Last Modified Time'] as string) || '',
      itemCode: ic,
      bomType: (r['BOM Item Type'] as string) || '',
      itemGroup: (r['Item Group'] as string) || '',
      itemName: (r['Item Name'] as string) || '',
      itemParent: (r['Item Parent'] as string) || '',
      convFactor: Number(r['Conversion Factor']) || 0,
      projUnits: round2(projUnits),
      customer: cust,
      projKg: round2(projKg),
      custGroup: cg,
      origin: origin,
      itemType: (r['Item Type'] as string) || '',
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
      isProj: true,
      isUnproj: false,
    });
  });

  // Add unprojected SO items
  Object.entries(soMap).forEach(([key, so]) => {
    if (projKeys.has(key)) return;
    const [cg, cust, origin, ic] = key.split('|');
    const soRow = soRaw.find(r => r['Item Code'] === ic);
    rows.push({
      month: '', year: '', lastModDate: '', lastModTime: '',
      itemCode: ic,
      bomType: '',
      itemGroup: (soRow?.['Item Group'] as string) || '',
      itemName: (soRow?.['Item Name'] as string) || '',
      itemParent: (soRow?.['Parent Item'] as string) || '',
      convFactor: Number(soRow?.['Conversion Factor']) || 0,
      projUnits: 0, customer: cust, projKg: 0,
      custGroup: cg, origin: origin,
      itemType: (soRow?.['Item Type'] as string) || '',
      newMIS: so.newMIS || '',
      dailyKg: 0, dailyUnits: 0, expKg: 0, expUnits: 0,
      soKg: round2(so.soKg), soUnits: round2(so.soUnits),
      diffKg: round2(-so.soKg), diffUnits: round2(-so.soUnits),
      achKg: 0, achUnits: 0,
      runRateKg: round2(daysElapsed > 0 ? (so.soKg / daysElapsed) * daysInMonth : 0),
      runRateUnits: round2(daysElapsed > 0 ? (so.soUnits / daysElapsed) * daysInMonth : 0),
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
    rows,
    summaryIG, summaryOrigin, summaryCG, summaryMIS,
    top20Kg, top20Units, bot20Kg, bot20Units,
  };
}
