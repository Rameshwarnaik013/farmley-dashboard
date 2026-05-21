import * as XLSX from 'xlsx';
import { DashData, Row, Summary } from './helpers';

// Safe divide helper
function div(a: number, b: number): number { return b === 0 ? 0 : a / b; }
// Round to N decimals
function rnd(v: number, d = 4): number { const m = Math.pow(10, d); return Math.round(v * m) / m; }

// ─── Config Sheet ───
function writeConfigSheet(wb: XLSX.WorkBook, data: DashData) {
  const cfg = data.config;
  const rows = [
    ['Run Rate Dashboard Configuration'],
    [],
    ['Parameter', 'Value'],
    ['Days in Month', cfg.daysInMonth],
    ['SO Date From', cfg.dateFrom],
    ['SO Date To', cfg.dateTo],
    ['Days Elapsed', cfg.daysElapsed],
    ['Total Rows', cfg.totalRows],
    ['SO Rows (Raw)', cfg.soRowsRaw],
    ['SO Rows (Clean)', cfg.soRowsClean],
  ];
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 20 }, { wch: 25 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Config');
}

// ─── Dashboard (main data) sheet ───
function writeDashboardSheet(wb: XLSX.WorkBook, data: DashData) {
  const headers = [
    'Customer Group', 'Customer', 'Origin', 'Item Code', 'Item Name', 'Item Group', 'New MIS',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units', 'MTD Ach%',
    'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units',
    'RunRate KGs', 'RunRate Units', 'Is Projected', 'Is Unprojected'
  ];
  const aoa: (string | number)[][] = [headers];

  data.rows.forEach(r => {
    aoa.push([
      r.custGroup, r.customer, r.origin, r.itemCode, r.itemName, r.itemGroup, r.newMIS,
      r.projKg, r.soKg, r.projUnits, rnd(r.expUnits, 0),
      rnd(r.expKg), r.soUnits, rnd(r.achKg),
      rnd(r.soVsProj), rnd(r.soLeftPct), rnd(r.soPctPerDay), rnd(r.daysToCover, 1),
      rnd(r.diffKg), rnd(r.diffUnits, 0),
      rnd(r.runRateKg), rnd(r.runRateUnits, 0),
      r.isProj ? 'Yes' : 'No',
      r.isUnproj ? 'Yes' : 'No',
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // Format percentage columns N(14), O(15), P(16), Q(17) — 1-indexed in headers
  const pctIdxs = [13, 14, 15, 16]; // 0-indexed: MTD Ach%, SO vs Proj%, % SO Left, SO%/Day
  for (let i = 0; i < data.rows.length; i++) {
    const row = i + 2;
    pctIdxs.forEach(c => {
      const addr = XLSX.utils.encode_cell({ r: row - 1, c });
      if (ws[addr]) ws[addr].z = '0.0%';
    });
  }
  ws['!cols'] = [
    { wch: 16 }, { wch: 20 }, { wch: 12 }, { wch: 18 }, { wch: 40 }, { wch: 16 }, { wch: 16 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 13 },
    { wch: 12 }, { wch: 12 }, { wch: 10 },
    { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 12 },
    { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 12 },
  ];
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.rows.length, c: headers.length - 1 } }) };
  XLSX.utils.book_append_sheet(wb, ws, 'Dashboard');
}

// ─── Grouped pivot sheet ───
function writeGroupedSheet(
  wb: XLSX.WorkBook, data: DashData,
  sheetName: string, groupKey: keyof Row, subGroupKey?: keyof Row
) {
  const DIM = data.config.daysInMonth;
  const DE = data.config.daysElapsed;

  const headers = [
    'Group', 'Sub Group', 'Item Code', 'Item Name',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units'
  ];
  const HC = headers.length;

  // Compute metrics from aggregated values
  function metrics(projKg: number, soKg: number, projUnits: number, soUnits: number) {
    const expSoUnits = DIM === 0 ? 0 : rnd((projUnits / DIM) * DE, 0);
    const expKg = DIM === 0 ? 0 : rnd((projKg / DIM) * DE);
    const achKg = rnd(div(soKg, expKg) * 100);
    const soVsProj = rnd(div(soKg, projKg) * 100);
    const soLeft = rnd(projKg > 0 ? 100 - soVsProj : 0);
    const soPctDay = rnd(DE > 0 ? div(soVsProj, DE) : 0);
    const daysLeft = rnd(soPctDay > 0 ? div(soLeft, soPctDay) : 0, 1);
    const diffKg = rnd(expKg - soKg);
    const diffUnits = rnd(expSoUnits - soUnits, 0);
    return [expSoUnits, expKg, soUnits, achKg, soVsProj, soLeft, soPctDay, daysLeft, diffKg, diffUnits];
  }

  // Build row from a single Row record
  function itemRow(g: string, sg: string, r: Row): (string | number)[] {
    const m = metrics(r.projKg, r.soKg, r.projUnits, r.soUnits);
    return [g, sg, r.itemCode, r.itemName, r.projKg, r.soKg, r.projUnits, ...m];
  }

  // Build aggregated row
  function aggRow(label: string, sg: string, info: string, rows: Row[]): (string | number)[] {
    const projKg = rows.reduce((s, r) => s + r.projKg, 0);
    const soKg = rows.reduce((s, r) => s + r.soKg, 0);
    const projUnits = rows.reduce((s, r) => s + r.projUnits, 0);
    const soUnits = rows.reduce((s, r) => s + r.soUnits, 0);
    const m = metrics(projKg, soKg, projUnits, soUnits);
    return [label, sg, '', info, rnd(projKg), rnd(soKg), rnd(projUnits, 0), ...m];
  }

  const aoa: (string | number)[][] = [headers];
  const outlineRows: { level: number; row: number }[] = [];

  // Group the rows
  const groups: Record<string, Row[]> = {};
  data.rows.forEach(r => {
    const g = (r[groupKey] as string) || 'Unknown';
    if (!groups[g]) groups[g] = [];
    groups[g].push(r);
  });
  const sortedGroups = Object.entries(groups).sort(([a], [b]) => a.localeCompare(b));

  sortedGroups.forEach(([gName, gRows]) => {
    const parentIdx = aoa.length;
    aoa.push(aggRow(gName, '', `(${gRows.length} items)`, gRows));

    if (subGroupKey) {
      const subGroups: Record<string, Row[]> = {};
      gRows.forEach(r => {
        const sg = (r[subGroupKey] as string) || 'Unknown';
        if (!subGroups[sg]) subGroups[sg] = [];
        subGroups[sg].push(r);
      });
      const sortedSubs = Object.entries(subGroups).sort(([a], [b]) => a.localeCompare(b));

      sortedSubs.forEach(([sgName, sgRows]) => {
        const subIdx = aoa.length;
        aoa.push(aggRow(gName, sgName, `(${sgRows.length})`, sgRows));
        outlineRows.push({ level: 1, row: subIdx });

        sgRows.forEach(r => {
          outlineRows.push({ level: 2, row: aoa.length });
          aoa.push(itemRow(gName, sgName, r));
        });
      });
    } else {
      gRows.forEach(r => {
        outlineRows.push({ level: 1, row: aoa.length });
        aoa.push(itemRow(gName, '', r));
      });
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Format percentage columns K(10), L(11), M(12), N(13)
  const pctIdxs = [10, 11, 12, 13];
  for (let i = 1; i < aoa.length; i++) {
    pctIdxs.forEach(c => {
      const addr = XLSX.utils.encode_cell({ r: i, c });
      if (ws[addr]) ws[addr].z = '0.0%';
    });
  }

  // Set row outlines
  if (!ws['!rows']) ws['!rows'] = [];
  outlineRows.forEach(({ level, row }) => {
    if (!ws['!rows']![row]) ws['!rows']![row] = {};
    ws['!rows']![row].level = level;
    ws['!rows']![row].hidden = true;
  });

  ws['!cols'] = [
    { wch: 30 }, { wch: 22 }, { wch: 18 }, { wch: 40 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 13 },
    { wch: 12 }, { wch: 12 },
    { wch: 10 }, { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 12 },
  ];
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 0, c: HC - 1 } }) };
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Summary sheet ───
function writeSummarySheet(wb: XLSX.WorkBook, summaries: Summary[], sheetName: string, data: DashData) {
  const headers = [
    'Name', 'Cust Count', 'Proj KGs', 'SO KGs', 'Proj Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left',
    'Diff KGs', 'Diff Units'
  ];
  const aoa: (string | number)[][] = [headers];

  let tCount = 0, tProjKg = 0, tSoKg = 0, tProjUnits = 0, tExpKg = 0, tSoUnits = 0;

  summaries.forEach(s => {
    const achKg = rnd(div(s.soKg, s.expKg) * 100);
    const soVsProj = rnd(div(s.soKg, s.projKg) * 100);
    const soLeft = rnd(s.projKg > 0 ? 100 - soVsProj : 0);
    aoa.push([
      s.name, s.count, rnd(s.projKg), rnd(s.soKg), rnd(s.projUnits, 0),
      rnd(s.expKg), rnd(s.soUnits, 0),
      achKg, soVsProj, soLeft,
      rnd(s.expKg - s.soKg), rnd(s.expUnits - s.soUnits, 0),
    ]);
    tCount += s.count; tProjKg += s.projKg; tSoKg += s.soKg;
    tProjUnits += s.projUnits; tExpKg += s.expKg; tSoUnits += s.soUnits;
  });

  // Totals row
  const tAch = rnd(div(tSoKg, tExpKg) * 100);
  const tSoVsProj = rnd(div(tSoKg, tProjKg) * 100);
  const tSoLeft = rnd(tProjKg > 0 ? 100 - tSoVsProj : 0);
  aoa.push([
    'TOTAL', tCount, rnd(tProjKg), rnd(tSoKg), rnd(tProjUnits, 0),
    rnd(tExpKg), rnd(tSoUnits, 0),
    tAch, tSoVsProj, tSoLeft,
    rnd(tExpKg - tSoKg), rnd(tProjUnits - tSoUnits, 0),
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // Format percentage columns H(7), I(8), J(9)
  const pctIdxs = [7, 8, 9];
  for (let i = 1; i < aoa.length; i++) {
    pctIdxs.forEach(c => {
      const addr = XLSX.utils.encode_cell({ r: i, c });
      if (ws[addr]) ws[addr].z = '0.0%';
    });
  }
  ws['!cols'] = [
    { wch: 28 }, { wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
    { wch: 14 }, { wch: 14 },
    { wch: 10 }, { wch: 11 }, { wch: 10 },
    { wch: 14 }, { wch: 14 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Leaderboard sheet ───
function writeLeaderboardSheet(wb: XLSX.WorkBook, rows: Row[], sheetName: string) {
  const headers = [
    'Rank', 'Item Code', 'Item Name', 'Cust Group', 'Customer', 'Origin',
    'Proj KGs', 'SO KGs', 'Exp KGs',
    'MTD Ach%', 'SO vs Proj%', '% SO Left',
    'SO%/Day', 'Days Left', 'Diff KGs', 'RunRate KGs'
  ];
  const aoa: (string | number)[][] = [headers];

  rows.forEach((r, i) => {
    aoa.push([
      i + 1, r.itemCode, r.itemName, r.custGroup, r.customer, r.origin,
      r.projKg, r.soKg, rnd(r.expKg),
      rnd(r.achKg), rnd(r.soVsProj), rnd(r.soLeftPct),
      rnd(r.soPctPerDay), rnd(r.daysToCover, 1), rnd(r.diffKg), rnd(r.runRateKg),
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // Format percentage columns J(9), K(10), L(11), M(12)
  const pctIdxs = [9, 10, 11, 12];
  for (let i = 0; i < rows.length; i++) {
    pctIdxs.forEach(c => {
      const addr = XLSX.utils.encode_cell({ r: i + 1, c });
      if (ws[addr]) ws[addr].z = '0.0%';
    });
  }
  ws['!cols'] = [
    { wch: 5 }, { wch: 18 }, { wch: 40 }, { wch: 16 }, { wch: 20 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 12 },
    { wch: 10 }, { wch: 11 }, { wch: 10 },
    { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 14 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Zero SO sheet ───
function writeZeroSOSheet(wb: XLSX.WorkBook, data: DashData) {
  const zeroRows = data.rows.filter(r => r.isProj && r.soKg === 0 && r.soUnits === 0 && r.projKg > 0);
  const headers = ['Item Code', 'Item Name', 'Customer Group', 'Customer', 'Origin', 'Item Group', 'Proj KGs', 'Proj Units'];
  const aoa: (string | number)[][] = [headers];
  zeroRows.forEach(r => {
    aoa.push([r.itemCode, r.itemName, r.custGroup, r.customer, r.origin, r.itemGroup, r.projKg, r.projUnits]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 18 }, { wch: 40 }, { wch: 16 }, { wch: 20 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Zero SO');
}

// ─── Unprojected SO sheet ───
function writeUnprojSheet(wb: XLSX.WorkBook, data: DashData) {
  const unprojRows = data.rows.filter(r => r.isUnproj);
  const headers = ['Item Code', 'Item Name', 'Customer Group', 'Customer', 'Origin', 'Item Group', 'SO KGs', 'SO Units'];
  const aoa: (string | number)[][] = [headers];
  unprojRows.forEach(r => {
    aoa.push([r.itemCode, r.itemName, r.custGroup, r.customer, r.origin, r.itemGroup, r.soKg, r.soUnits]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 18 }, { wch: 40 }, { wch: 16 }, { wch: 20 }, { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Unprojected SO');
}

// ─── CFA Items sheet ───
function writeCFASheet(wb: XLSX.WorkBook, data: DashData, cfaNames: Set<string>) {
  const cfaRows = data.rows.filter(r => cfaNames.has(r.itemName));
  const DIM = data.config.daysInMonth;
  const DE = data.config.daysElapsed;
  const headers = [
    'Item Name', 'Customer', 'Item Code', 'Origin',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units'
  ];
  const aoa: (string | number)[][] = [headers];

  let tProjKg = 0, tSoKg = 0, tProjUnits = 0, tSoUnits = 0;

  cfaRows.forEach(r => {
    const expSoUnits = DIM === 0 ? 0 : rnd((r.projUnits / DIM) * DE, 0);
    const expKg = DIM === 0 ? 0 : rnd((r.projKg / DIM) * DE);
    const achKg = rnd(div(r.soKg, expKg) * 100);
    const soVsProj = rnd(div(r.soKg, r.projKg) * 100);
    const soLeft = rnd(r.projKg > 0 ? 100 - soVsProj : 0);
    const soPctDay = rnd(DE > 0 ? div(soVsProj, DE) : 0);
    const daysLeft = rnd(soPctDay > 0 ? div(soLeft, soPctDay) : 0, 1);

    aoa.push([
      r.itemName, r.customer, r.itemCode, r.origin,
      r.projKg, r.soKg, r.projUnits, expSoUnits,
      expKg, r.soUnits,
      achKg, soVsProj, soLeft, soPctDay, daysLeft,
      rnd(expKg - r.soKg), rnd(expSoUnits - r.soUnits, 0),
    ]);
    tProjKg += r.projKg; tSoKg += r.soKg; tProjUnits += r.projUnits; tSoUnits += r.soUnits;
  });

  // Totals
  const tExpSoUnits = DIM === 0 ? 0 : rnd((tProjUnits / DIM) * DE, 0);
  const tExpKg = DIM === 0 ? 0 : rnd((tProjKg / DIM) * DE);
  const tAch = rnd(div(tSoKg, tExpKg) * 100);
  const tSoVsProj = rnd(div(tSoKg, tProjKg) * 100);
  const tSoLeft = rnd(tProjKg > 0 ? 100 - tSoVsProj : 0);
  const tSoPctDay = rnd(DE > 0 ? div(tSoVsProj, DE) : 0);
  const tDaysLeft = rnd(tSoPctDay > 0 ? div(tSoLeft, tSoPctDay) : 0, 1);

  aoa.push([
    'TOTAL', '', '', '',
    rnd(tProjKg), rnd(tSoKg), rnd(tProjUnits, 0), tExpSoUnits,
    tExpKg, rnd(tSoUnits, 0),
    tAch, tSoVsProj, tSoLeft, tSoPctDay, tDaysLeft,
    rnd(tExpKg - tSoKg), rnd(tExpSoUnits - tSoUnits, 0),
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // Format percentage columns K(10), L(11), M(12), N(13)
  const pctIdxs = [10, 11, 12, 13];
  for (let i = 1; i < aoa.length; i++) {
    pctIdxs.forEach(c => {
      const addr = XLSX.utils.encode_cell({ r: i, c });
      if (ws[addr]) ws[addr].z = '0.0%';
    });
  }
  ws['!cols'] = [
    { wch: 45 }, { wch: 20 }, { wch: 18 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 13 },
    { wch: 12 }, { wch: 12 },
    { wch: 10 }, { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 12 },
  ];
  ws['!autofilter'] = { ref: 'A1:Q1' };
  XLSX.utils.book_append_sheet(wb, ws, 'CFA Items');
}

// ─── Main export function ───
export function exportDashboardExcel(data: DashData, cfaNames: Set<string>) {
  const wb = XLSX.utils.book_new();

  writeConfigSheet(wb, data);
  writeDashboardSheet(wb, data);
  writeGroupedSheet(wb, data, 'By New MIS', 'newMIS');
  writeGroupedSheet(wb, data, 'By Customer Group', 'custGroup', 'customer');
  writeGroupedSheet(wb, data, 'By Origin', 'origin');
  writeGroupedSheet(wb, data, 'By Item Name', 'itemName', 'customer');
  writeCFASheet(wb, data, cfaNames);
  writeSummarySheet(wb, data.summaryIG, 'Summary - Item Group', data);
  writeSummarySheet(wb, data.summaryOrigin, 'Summary - Origin', data);
  writeLeaderboardSheet(wb, data.top20Kg, 'Top 20 KGs');
  writeLeaderboardSheet(wb, data.top20Units, 'Top 20 Units');
  writeLeaderboardSheet(wb, data.bot20Kg, 'Bottom 20 KGs');
  writeLeaderboardSheet(wb, data.bot20Units, 'Bottom 20 Units');
  writeZeroSOSheet(wb, data);
  writeUnprojSheet(wb, data);

  const filename = `RunRate_Dashboard_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, filename);
}
