import * as XLSX from 'xlsx';
import { DashData, Row, Summary } from './helpers';

// Column letters helper
function colLetter(n: number): string {
  let s = '';
  n++;
  while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
  return s;
}

function applyHeaderStyle(ws: XLSX.WorkSheet, cols: number, row: number) {
  for (let c = 0; c < cols; c++) {
    const addr = colLetter(c) + row;
    if (!ws[addr]) ws[addr] = { v: '' };
    ws[addr].s = { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '4472C4' } }, alignment: { horizontal: 'center' } };
  }
}

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

// ─── Dashboard (main data) sheet with formulas ───
function writeDashboardSheet(wb: XLSX.WorkBook, data: DashData) {
  const headers = [
    'Customer Group', 'Customer', 'Origin', 'Item Code', 'Item Name', 'Item Group', 'New MIS',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units', 'MTD Ach%',
    'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units',
    'RunRate KGs', 'RunRate Units', 'Is Projected', 'Is Unprojected'
  ];
  // Config references: B4=Days in Month, B7=Days Elapsed
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';

  // Column mapping: H=ProjKg, I=SOKg, J=ProjUnits, K=ExpSOUnits, L=ExpKg, M=SOUnits
  // N=MTDAch%, O=SOvsProj%, P=%SOLeft, Q=SO%/Day, R=DaysLeft
  // S=DiffKg, T=DiffUnits, U=RunRateKg, V=RunRateUnits
  const aoa: unknown[][] = [headers];

  data.rows.forEach((r, i) => {
    const row = i + 2;
    aoa.push([
      r.custGroup, r.customer, r.origin, r.itemCode, r.itemName, r.itemGroup, r.newMIS,
      r.projKg,                                        // H: Proj KGs
      r.soKg,                                          // I: SO KGs
      r.projUnits,                                     // J: Proj Units
      { f: `IF(${DIM}=0,0,(J${row}/${DIM})*${DE})` }, // K: Exp SO Units = (ProjUnits/DIM)*DE
      { f: `IF(${DIM}=0,0,(H${row}/${DIM})*${DE})` }, // L: Exp KGs = (ProjKg/DIM)*DE
      r.soUnits,                                       // M: SO Units
      { f: `IF(L${row}=0,0,I${row}/L${row})` },       // N: MTD Ach% = SOKg/ExpKg
      { f: `IF(H${row}=0,0,I${row}/H${row})` },       // O: SO vs Proj% = SOKg/ProjKg
      { f: `IF(H${row}=0,0,1-O${row})` },             // P: % SO Left = 1 - SOvsProj
      { f: `IF(${DE}=0,0,O${row}/${DE})` },            // Q: SO%/Day = SOvsProj/DaysElapsed
      { f: `IF(Q${row}=0,0,P${row}/Q${row})` },       // R: Days Left = %SOLeft/SO%perDay
      { f: `L${row}-I${row}` },                        // S: Diff KGs = ExpKg - SOKg
      { f: `K${row}-M${row}` },                        // T: Diff Units = ExpSOUnits - SOUnits
      { f: `IF(${DE}=0,0,(I${row}/${DE})*${DIM})` },  // U: RunRate KGs
      { f: `IF(${DE}=0,0,(M${row}/${DE})*${DIM})` },  // V: RunRate Units
      r.isProj ? 'Yes' : 'No',
      r.isUnproj ? 'Yes' : 'No',
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Format percentage columns
  const pctCols = ['N', 'O', 'P', 'Q']; // MTD Ach%, SO vs Proj%, % SO Left, SO%/Day
  for (let i = 0; i < data.rows.length; i++) {
    const row = i + 2;
    pctCols.forEach(c => { if (ws[`${c}${row}`]) ws[`${c}${row}`].z = '0.0%'; });
  }

  ws['!cols'] = [
    { wch: 16 }, { wch: 20 }, { wch: 12 }, { wch: 18 }, { wch: 40 }, { wch: 16 }, { wch: 16 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 13 },
    { wch: 12 }, { wch: 12 }, { wch: 10 },
    { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 12 },
    { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 12 },
  ];

  ws['!autofilter'] = { ref: `A1:${colLetter(headers.length - 1)}${data.rows.length + 1}` };
  XLSX.utils.book_append_sheet(wb, ws, 'Dashboard');
}

// ─── Grouped pivot sheet with formulas ───
function writeGroupedSheet(
  wb: XLSX.WorkBook, data: DashData,
  sheetName: string, groupKey: keyof Row, subGroupKey?: keyof Row
) {
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';

  const headers = [
    'Group', 'Sub Group', 'Item Code', 'Item Name',
    'Proj KGs', 'Proj Units', 'Daily Demand KG', 'Daily Demand Units',
    'Expected KGs', 'Expected Units', 'SO KGs', 'SO Units',
    'Diff KGs', 'Diff Units', 'Ach% KGs', 'Ach% Units'
  ];

  const aoa: unknown[][] = [headers];

  // Group the rows
  const groups: Record<string, Row[]> = {};
  data.rows.forEach(r => {
    const g = (r[groupKey] as string) || 'Unknown';
    if (!groups[g]) groups[g] = [];
    groups[g].push(r);
  });

  const sortedGroups = Object.entries(groups).sort(([a], [b]) => a.localeCompare(b));
  const outlineRows: { level: number; row: number }[] = [];

  sortedGroups.forEach(([gName, gRows]) => {
    // Parent row (aggregation with SUBTOTAL formulas added after children)
    const parentRowIdx = aoa.length; // 0-indexed in aoa

    if (subGroupKey) {
      // Sub-group within parent
      const subGroups: Record<string, Row[]> = {};
      gRows.forEach(r => {
        const sg = (r[subGroupKey] as string) || 'Unknown';
        if (!subGroups[sg]) subGroups[sg] = [];
        subGroups[sg].push(r);
      });

      const sortedSubs = Object.entries(subGroups).sort(([a], [b]) => a.localeCompare(b));

      // Parent placeholder
      aoa.push([gName, '', '', `(${gRows.length} items)`, 0, 0, null, null, null, null, 0, 0, null, null, null, null]);
      const parentExcelRow = aoa.length; // 1-indexed

      const childStartRows: number[] = [];

      sortedSubs.forEach(([sgName, sgRows]) => {
        // Sub-group summary row
        aoa.push([gName, sgName, '', `(${sgRows.length})`, 0, 0, null, null, null, null, 0, 0, null, null, null, null]);
        const subParentExcelRow = aoa.length;
        const subChildStart = aoa.length + 1;

        sgRows.forEach(r => {
          const exRow = aoa.length + 1;
          aoa.push([
            gName, sgName, r.itemCode, r.itemName,
            r.projKg, r.projUnits,
            { f: `E${exRow}/${DIM}` }, { f: `F${exRow}/${DIM}` },
            { f: `G${exRow}*${DE}` }, { f: `H${exRow}*${DE}` },
            r.soKg, r.soUnits,
            { f: `I${exRow}-K${exRow}` }, { f: `J${exRow}-L${exRow}` },
            { f: `IF(I${exRow}=0,0,K${exRow}/I${exRow})` },
            { f: `IF(J${exRow}=0,0,L${exRow}/J${exRow})` },
          ]);
          outlineRows.push({ level: 2, row: aoa.length - 1 });
        });

        const subChildEnd = aoa.length;
        // Fill sub-group summary with SUM formulas
        const sgr = aoa[subParentExcelRow - 1] as unknown[];
        sgr[4] = { f: `SUM(E${subChildStart}:E${subChildEnd})` };
        sgr[5] = { f: `SUM(F${subChildStart}:F${subChildEnd})` };
        sgr[6] = { f: `${colLetter(4)}${subParentExcelRow}/${DIM}` };
        sgr[7] = { f: `${colLetter(5)}${subParentExcelRow}/${DIM}` };
        sgr[8] = { f: `${colLetter(6)}${subParentExcelRow}*${DE}` };
        sgr[9] = { f: `${colLetter(7)}${subParentExcelRow}*${DE}` };
        sgr[10] = { f: `SUM(K${subChildStart}:K${subChildEnd})` };
        sgr[11] = { f: `SUM(L${subChildStart}:L${subChildEnd})` };
        sgr[12] = { f: `I${subParentExcelRow}-K${subParentExcelRow}` };
        sgr[13] = { f: `J${subParentExcelRow}-L${subParentExcelRow}` };
        sgr[14] = { f: `IF(I${subParentExcelRow}=0,0,K${subParentExcelRow}/I${subParentExcelRow})` };
        sgr[15] = { f: `IF(J${subParentExcelRow}=0,0,L${subParentExcelRow}/J${subParentExcelRow})` };

        childStartRows.push(subParentExcelRow);
        outlineRows.push({ level: 1, row: subParentExcelRow - 1 });
      });

      // Fill parent summary
      const pr = aoa[parentRowIdx] as unknown[];
      const sumParts = (col: string) => childStartRows.map(r => `${col}${r}`).join('+');
      pr[4] = { f: sumParts('E') };
      pr[5] = { f: sumParts('F') };
      pr[6] = { f: `E${parentExcelRow}/${DIM}` };
      pr[7] = { f: `F${parentExcelRow}/${DIM}` };
      pr[8] = { f: `G${parentExcelRow}*${DE}` };
      pr[9] = { f: `H${parentExcelRow}*${DE}` };
      pr[10] = { f: sumParts('K') };
      pr[11] = { f: sumParts('L') };
      pr[12] = { f: `I${parentExcelRow}-K${parentExcelRow}` };
      pr[13] = { f: `J${parentExcelRow}-L${parentExcelRow}` };
      pr[14] = { f: `IF(I${parentExcelRow}=0,0,K${parentExcelRow}/I${parentExcelRow})` };
      pr[15] = { f: `IF(J${parentExcelRow}=0,0,L${parentExcelRow}/J${parentExcelRow})` };

    } else {
      // No sub-groups, just parent + child items
      aoa.push([gName, '', '', `(${gRows.length} items)`, 0, 0, null, null, null, null, 0, 0, null, null, null, null]);
      const parentExcelRow = aoa.length;
      const childStart = aoa.length + 1;

      gRows.forEach(r => {
        const exRow = aoa.length + 1;
        aoa.push([
          gName, '', r.itemCode, r.itemName,
          r.projKg, r.projUnits,
          { f: `E${exRow}/${DIM}` }, { f: `F${exRow}/${DIM}` },
          { f: `G${exRow}*${DE}` }, { f: `H${exRow}*${DE}` },
          r.soKg, r.soUnits,
          { f: `I${exRow}-K${exRow}` }, { f: `J${exRow}-L${exRow}` },
          { f: `IF(I${exRow}=0,0,K${exRow}/I${exRow})` },
          { f: `IF(J${exRow}=0,0,L${exRow}/J${exRow})` },
        ]);
        outlineRows.push({ level: 1, row: aoa.length - 1 });
      });

      const childEnd = aoa.length;
      const pr = aoa[parentRowIdx] as unknown[];
      pr[4] = { f: `SUM(E${childStart}:E${childEnd})` };
      pr[5] = { f: `SUM(F${childStart}:F${childEnd})` };
      pr[6] = { f: `E${parentExcelRow}/${DIM}` };
      pr[7] = { f: `F${parentExcelRow}/${DIM}` };
      pr[8] = { f: `G${parentExcelRow}*${DE}` };
      pr[9] = { f: `H${parentExcelRow}*${DE}` };
      pr[10] = { f: `SUM(K${childStart}:K${childEnd})` };
      pr[11] = { f: `SUM(L${childStart}:L${childEnd})` };
      pr[12] = { f: `I${parentExcelRow}-K${parentExcelRow}` };
      pr[13] = { f: `J${parentExcelRow}-L${parentExcelRow}` };
      pr[14] = { f: `IF(I${parentExcelRow}=0,0,K${parentExcelRow}/I${parentExcelRow})` };
      pr[15] = { f: `IF(J${parentExcelRow}=0,0,L${parentExcelRow}/J${parentExcelRow})` };
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Format Ach% columns as percentage
  for (let i = 1; i < aoa.length; i++) {
    const row = i + 1;
    const cellO = ws[`O${row}`];
    const cellP = ws[`P${row}`];
    if (cellO) cellO.z = '0.0%';
    if (cellP) cellP.z = '0.0%';
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
    { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 14 },
    { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 10 },
  ];

  ws['!autofilter'] = { ref: `A1:P1` };
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Summary sheet ───
function writeSummarySheet(wb: XLSX.WorkBook, summaries: Summary[], sheetName: string) {
  const headers = ['Name', 'Count', 'Proj KGs', 'Proj Units', 'Exp KGs', 'Exp Units', 'SO KGs', 'SO Units', 'Diff KGs', 'Diff Units', 'Ach% KGs', 'Ach% Units'];
  const aoa: unknown[][] = [headers];

  summaries.forEach(s => {
    const row = aoa.length + 1;
    aoa.push([
      s.name, s.count, s.projKg, s.projUnits, s.expKg, s.expUnits, s.soKg, s.soUnits,
      { f: `E${row}-G${row}` },    // Diff KGs = Exp - SO
      { f: `F${row}-H${row}` },    // Diff Units
      { f: `IF(E${row}=0,0,G${row}/E${row})` },  // Ach% KGs
      { f: `IF(F${row}=0,0,H${row}/F${row})` },  // Ach% Units
    ]);
  });

  // Totals row
  const totalRow = aoa.length + 1;
  const lastData = aoa.length;
  aoa.push([
    'TOTAL',
    { f: `SUM(B2:B${lastData})` },
    { f: `SUM(C2:C${lastData})` }, { f: `SUM(D2:D${lastData})` },
    { f: `SUM(E2:E${lastData})` }, { f: `SUM(F2:F${lastData})` },
    { f: `SUM(G2:G${lastData})` }, { f: `SUM(H2:H${lastData})` },
    { f: `E${totalRow}-G${totalRow}` }, { f: `F${totalRow}-H${totalRow}` },
    { f: `IF(E${totalRow}=0,0,G${totalRow}/E${totalRow})` },
    { f: `IF(F${totalRow}=0,0,H${totalRow}/F${totalRow})` },
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Format Ach% as percentage
  for (let i = 1; i <= summaries.length + 1; i++) {
    const row = i + 1;
    if (ws[`K${row}`]) ws[`K${row}`].z = '0.0%';
    if (ws[`L${row}`]) ws[`L${row}`].z = '0.0%';
  }

  ws['!cols'] = [
    { wch: 28 }, { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
    { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 10 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Leaderboard sheet ───
function writeLeaderboardSheet(wb: XLSX.WorkBook, rows: Row[], sheetName: string) {
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';
  const headers = [
    'Rank', 'Item Code', 'Item Name', 'Customer Group', 'Customer', 'Origin',
    'Proj KGs', 'Proj Units', 'Daily KG', 'Daily Units',
    'Expected KGs', 'Expected Units', 'SO KGs', 'SO Units',
    'Diff KGs', 'Ach% KGs', 'Ach% Units'
  ];
  const aoa: unknown[][] = [headers];

  rows.forEach((r, i) => {
    const row = i + 2;
    aoa.push([
      i + 1, r.itemCode, r.itemName, r.custGroup, r.customer, r.origin,
      r.projKg, r.projUnits,
      { f: `G${row}/${DIM}` }, { f: `H${row}/${DIM}` },
      { f: `I${row}*${DE}` }, { f: `J${row}*${DE}` },
      r.soKg, r.soUnits,
      { f: `K${row}-M${row}` },
      { f: `IF(K${row}=0,0,M${row}/K${row})` },
      { f: `IF(L${row}=0,0,N${row}/L${row})` },
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  for (let i = 0; i < rows.length; i++) {
    const row = i + 2;
    if (ws[`P${row}`]) ws[`P${row}`].z = '0.0%';
    if (ws[`Q${row}`]) ws[`Q${row}`].z = '0.0%';
  }

  ws['!cols'] = [
    { wch: 5 }, { wch: 18 }, { wch: 40 }, { wch: 16 }, { wch: 20 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 10 }, { wch: 10 },
  ];

  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Zero SO sheet ───
function writeZeroSOSheet(wb: XLSX.WorkBook, data: DashData) {
  const zeroRows = data.rows.filter(r => r.isProj && r.soKg === 0 && r.soUnits === 0 && r.projKg > 0);
  const headers = ['Item Code', 'Item Name', 'Customer Group', 'Customer', 'Origin', 'Item Group', 'Proj KGs', 'Proj Units'];
  const aoa: unknown[][] = [headers];
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
  const aoa: unknown[][] = [headers];
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
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';
  const headers = [
    'Item Name', 'Customer', 'Item Code', 'Origin',
    'Proj KGs', 'Proj Units', 'Daily KG', 'Daily Units',
    'Expected KGs', 'Expected Units', 'SO KGs', 'SO Units',
    'Diff KGs', 'Diff Units', 'Ach% KGs', 'Ach% Units'
  ];
  const aoa: unknown[][] = [headers];
  cfaRows.forEach(r => {
    const row = aoa.length + 1;
    aoa.push([
      r.itemName, r.customer, r.itemCode, r.origin,
      r.projKg, r.projUnits,
      { f: `E${row}/${DIM}` }, { f: `F${row}/${DIM}` },
      { f: `G${row}*${DE}` }, { f: `H${row}*${DE}` },
      r.soKg, r.soUnits,
      { f: `I${row}-K${row}` }, { f: `J${row}-L${row}` },
      { f: `IF(I${row}=0,0,K${row}/I${row})` },
      { f: `IF(J${row}=0,0,L${row}/J${row})` },
    ]);
  });

  // Totals
  const totalRow = aoa.length + 1;
  const last = aoa.length;
  aoa.push([
    'TOTAL', '', '', '',
    { f: `SUM(E2:E${last})` }, { f: `SUM(F2:F${last})` },
    null, null,
    { f: `SUM(I2:I${last})` }, { f: `SUM(J2:J${last})` },
    { f: `SUM(K2:K${last})` }, { f: `SUM(L2:L${last})` },
    { f: `I${totalRow}-K${totalRow}` }, { f: `J${totalRow}-L${totalRow}` },
    { f: `IF(I${totalRow}=0,0,K${totalRow}/I${totalRow})` },
    { f: `IF(J${totalRow}=0,0,L${totalRow}/J${totalRow})` },
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (let i = 1; i < aoa.length; i++) {
    const row = i + 1;
    if (ws[`O${row}`]) ws[`O${row}`].z = '0.0%';
    if (ws[`P${row}`]) ws[`P${row}`].z = '0.0%';
  }
  ws['!cols'] = [
    { wch: 45 }, { wch: 20 }, { wch: 18 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    { wch: 14 }, { wch: 14 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 10 },
  ];
  ws['!autofilter'] = { ref: `A1:P1` };
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
  writeSummarySheet(wb, data.summaryIG, 'Summary - Item Group');
  writeSummarySheet(wb, data.summaryOrigin, 'Summary - Origin');
  writeLeaderboardSheet(wb, data.top20Kg, 'Top 20 KGs');
  writeLeaderboardSheet(wb, data.top20Units, 'Top 20 Units');
  writeLeaderboardSheet(wb, data.bot20Kg, 'Bottom 20 KGs');
  writeLeaderboardSheet(wb, data.bot20Units, 'Bottom 20 Units');
  writeZeroSOSheet(wb, data);
  writeUnprojSheet(wb, data);

  const filename = `RunRate_Dashboard_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, filename);
}
