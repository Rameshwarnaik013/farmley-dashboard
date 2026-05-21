import XLSX from 'xlsx';
import { writeFileSync } from 'fs';
import { join, dirname, basename } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJ_FILE = join(__dirname, '..', '..', 'Downloads', 'May 2026 Sales Projection (5).xlsx');
const SO_FILE = join(__dirname, '..', '..', 'Downloads', 'Custom Sales Order Report (40).xlsx');
const OUT = join(__dirname, '..', 'public', 'data.json');
const DAYS_IN_MONTH = 30;
const VALID_IT = ['ashish.k@farmley.com', 'abhishek.ku@farmley.com'];

function readExcel(path) {
  const wb = XLSX.readFile(path);
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
}

function cleanSO(rows) {
  return rows.filter(r => {
    if (r['Sales Order Created By'] === 'Administrator') return false;
    if (r['Workflow State'] === 'Internal Transfer' && !VALID_IT.includes(r['Sales Order Created By'])) return false;
    if (['On Hold', 'Rejected', 'Cancelled'].includes(r['Workflow State'])) return false;
    if (r['Purpose'] === 'Documentation') return false;
    if (['Cancelled', 'On Hold', 'Rejected'].includes(r['Sales Order Status'])) return false;
    if (r['Sales Order Status'] === 'Internal Transfer' && !VALID_IT.includes(r['Sales Order Created By'])) return false;
    if (typeof r['Customer'] === 'string' && r['Customer'].startsWith('Sample Order')) return false;
    if (r['Item Type'] === 'Bulk') return false;
    if (r['Returned Qty'] > 1) return false;
    return true;
  });
}

function parseDate(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number') {
    const d = new Date((v - 25569) * 86400000);
    return d;
  }
  return new Date(v);
}

console.log('Reading files...');
const projRaw = readExcel(PROJ_FILE);
let soRaw = readExcel(SO_FILE);
const soRawCount = soRaw.length;

console.log(`  Projections: ${projRaw.length} rows`);
console.log(`  Sales Orders: ${soRaw.length} rows`);

console.log('Cleaning SO...');
soRaw = cleanSO(soRaw);
console.log(`  After cleaning: ${soRaw.length} rows`);

// Remap
projRaw.forEach(r => {
  if (r['Customer Group'] === 'Category A' && r['Customer'] === 'Jhabak Marketing')
    r['Customer Group'] = 'Modern Trade';
});
soRaw.forEach(r => {
  if (r['Customer Group'] === 'Category A') r['Customer'] = 'GT Retail';
});

// Date range — always from 1st of month to max SO date
const soDates = soRaw.map(r => parseDate(r['Sales Order Date'])).filter(d => !isNaN(d));
soDates.sort((a, b) => a - b);
const dateTo = soDates[soDates.length - 1];
const dateToDay = dateTo.getUTCDate() || dateTo.getDate();
const dateToMonth = dateTo.getUTCMonth() ?? dateTo.getMonth();
const dateToYear = dateTo.getUTCFullYear() || dateTo.getFullYear();
const dateFrom = new Date(Date.UTC(dateToYear, dateToMonth, 1)); // 1st of month UTC
const daysElapsed = dateToDay; // day of month = days from 1st to max SO date
const dateFromStr = dateFrom.toISOString().slice(0, 10);
const dateToStr = dateTo.toISOString().slice(0, 10);
console.log(`  Date range: ${dateFromStr} to ${dateToStr}, ${daysElapsed} days (1st to ${dateToDay}th)`);

// Build SO lookup: key = CG|Customer|Origin|ItemCode -> {soKg, soUnits, newMIS}
const soMap = {};
soRaw.forEach(r => {
  const key = `${r['Customer Group']}|${r['Customer']}|${r['Origin']}|${r['Item Code']}`;
  if (!soMap[key]) soMap[key] = { soKg: 0, soUnits: 0, newMIS: r['NEW MIS ITEM GROUP'] || '' };
  soMap[key].soKg += (Number(r['Stock Qty']) || 0);
  soMap[key].soUnits += (Number(r['Qty']) || 0);
  if (!soMap[key].newMIS && r['NEW MIS ITEM GROUP']) soMap[key].newMIS = r['NEW MIS ITEM GROUP'];
});

// Also build SO lookup by just ItemCode for New MIS mapping
const soItemMIS = {};
soRaw.forEach(r => {
  const ic = r['Item Code'];
  if (r['NEW MIS ITEM GROUP'] && !soItemMIS[ic]) soItemMIS[ic] = r['NEW MIS ITEM GROUP'];
});

// Build enriched projection rows
const rows = [];
const projKeys = new Set();

projRaw.forEach(r => {
  const cg = r['Customer Group'] || '';
  const cust = r['Customer'] || '';
  const origin = r['Origin'] || '';
  const ic = r['Item Code'] || '';
  const key = `${cg}|${cust}|${origin}|${ic}`;
  projKeys.add(key);

  const projKg = Number(r['Total KGs']) || 0;
  const projUnits = Number(r['Projection Units']) || 0;
  const dailyKg = projKg / DAYS_IN_MONTH;
  const dailyUnits = projUnits / DAYS_IN_MONTH;
  const expKg = dailyKg * daysElapsed;
  const expUnits = dailyUnits * daysElapsed;

  const so = soMap[key] || { soKg: 0, soUnits: 0, newMIS: '' };
  const achKg = expKg > 0 ? (so.soKg / expKg) * 100 : 0;
  const achUnits = expUnits > 0 ? (so.soUnits / expUnits) * 100 : 0;

  rows.push({
    month: r['Month'] || '',
    year: r['Year'] || '',
    lastModDate: r['Last Modified Date'] || '',
    lastModTime: r['Last Modified Time'] || '',
    itemCode: ic,
    bomType: r['BOM Item Type'] || '',
    itemGroup: r['Item Group'] || '',
    itemName: r['Item Name'] || '',
    itemParent: r['Item Parent'] || '',
    convFactor: Number(r['Conversion Factor']) || 0,
    projUnits: round2(projUnits),
    customer: cust,
    projKg: round2(projKg),
    custGroup: cg,
    origin: origin,
    itemType: r['Item Type'] || '',
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
    runRateKg: round2(daysElapsed > 0 ? (so.soKg / daysElapsed) * DAYS_IN_MONTH : 0),
    runRateUnits: round2(daysElapsed > 0 ? (so.soUnits / daysElapsed) * DAYS_IN_MONTH : 0),
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
    itemGroup: soRow?.['Item Group'] || '',
    itemName: soRow?.['Item Name'] || '',
    itemParent: soRow?.['Parent Item'] || '',
    convFactor: Number(soRow?.['Conversion Factor']) || 0,
    projUnits: 0, customer: cust, projKg: 0,
    custGroup: cg, origin: origin,
    itemType: soRow?.['Item Type'] || '',
    newMIS: so.newMIS || '',
    dailyKg: 0, dailyUnits: 0, expKg: 0, expUnits: 0,
    soKg: round2(so.soKg), soUnits: round2(so.soUnits),
    diffKg: round2(-so.soKg), diffUnits: round2(-so.soUnits),
    achKg: 0, achUnits: 0,
    runRateKg: round2(daysElapsed > 0 ? (so.soKg / daysElapsed) * DAYS_IN_MONTH : 0),
    runRateUnits: round2(daysElapsed > 0 ? (so.soUnits / daysElapsed) * DAYS_IN_MONTH : 0),
    isProj: false, isUnproj: true,
  });
});

// Pre-aggregate summaries
function aggGroup(rows, groupKey) {
  const map = {};
  rows.forEach(r => {
    const g = r[groupKey] || 'Unknown';
    if (!map[g]) map[g] = { name: g, projKg: 0, projUnits: 0, expKg: 0, expUnits: 0, soKg: 0, soUnits: 0, count: 0 };
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

// Collect filter values
const unique = (arr, key) => [...new Set(arr.map(r => r[key]).filter(Boolean))].sort();

const output = {
  config: {
    daysInMonth: DAYS_IN_MONTH,
    dateFrom: dateFromStr,
    dateTo: dateToStr,
    daysElapsed,
    soRowsRaw: soRawCount,
    soRowsClean: soRaw.length,
    totalRows: rows.length,
    projFile: basename(PROJ_FILE),
    soFile: basename(SO_FILE),
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

writeFileSync(OUT, JSON.stringify(output));
console.log(`\nOutput: ${OUT} (${(JSON.stringify(output).length / 1024 / 1024).toFixed(1)} MB)`);
console.log(`  Total rows: ${rows.length}`);
console.log(`  Projected rows: ${rows.filter(r => r.isProj).length}`);
console.log(`  Unprojected SO rows: ${rows.filter(r => r.isUnproj).length}`);

function round2(v) { return Math.round(v * 100) / 100; }
