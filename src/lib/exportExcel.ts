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
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';
  // H=ProjKg I=SOKg J=ProjUnits K=ExpSOUnits L=ExpKg M=SOUnits
  // N=MTDAch% O=SOvsProj% P=%SOLeft Q=SO%/Day R=DaysLeft
  // S=DiffKg T=DiffUnits U=RunRateKg V=RunRateUnits
  const aoa: unknown[][] = [headers];

  data.rows.forEach((r, i) => {
    const row = i + 2;
    aoa.push([
      r.custGroup, r.customer, r.origin, r.itemCode, r.itemName, r.itemGroup, r.newMIS,
      r.projKg,                                        // H
      r.soKg,                                          // I
      r.projUnits,                                     // J
      { f: `IF(${DIM}=0,0,(J${row}/${DIM})*${DE})` }, // K: Exp SO Units
      { f: `IF(${DIM}=0,0,(H${row}/${DIM})*${DE})` }, // L: Exp KGs
      r.soUnits,                                       // M
      { f: `IF(L${row}=0,0,I${row}/L${row})` },       // N: MTD Ach%
      { f: `IF(H${row}=0,0,I${row}/H${row})` },       // O: SO vs Proj%
      { f: `IF(H${row}=0,0,1-O${row})` },             // P: % SO Left
      { f: `IF(${DE}=0,0,O${row}/${DE})` },            // Q: SO%/Day
      { f: `IF(Q${row}=0,0,P${row}/Q${row})` },       // R: Days Left
      { f: `L${row}-I${row}` },                        // S: Diff KGs
      { f: `K${row}-M${row}` },                        // T: Diff Units
      { f: `IF(${DE}=0,0,(I${row}/${DE})*${DIM})` },  // U: RunRate KGs
      { f: `IF(${DE}=0,0,(M${row}/${DE})*${DIM})` },  // V: RunRate Units
      r.isProj ? 'Yes' : 'No',
      r.isUnproj ? 'Yes' : 'No',
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const pctCols = ['N', 'O', 'P', 'Q'];
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
// New column layout:
// A=Group B=SubGroup C=ItemCode D=ItemName
// E=ProjKGs F=SOKGs G=ProjUnits H=ExpSOUnits I=ExpKGs J=SOUnits
// K=MTDAch% L=SOvsProj% M=%SOLeft N=SO%/Day O=DaysLeft
// P=DiffKGs Q=DiffUnits
function writeGroupedSheet(
  wb: XLSX.WorkBook, data: DashData,
  sheetName: string, groupKey: keyof Row, subGroupKey?: keyof Row
) {
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';

  const headers = [
    'Group', 'Sub Group', 'Item Code', 'Item Name',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units'
  ];
  const HC = headers.length; // 17 columns

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

  // Helper: build formula cells for a row given the Excel row number
  // Expects E=ProjKg(val), F=SOKg(val), G=ProjUnits(val), J=SOUnits(val) to already be values/SUM
  function formulaCells(r: number): Record<number, { f: string }> {
    return {
      7:  { f: `IF(${DIM}=0,0,(G${r}/${DIM})*${DE})` },              // H: Exp SO Units
      8:  { f: `IF(${DIM}=0,0,(E${r}/${DIM})*${DE})` },              // I: Exp KGs
      10: { f: `IF(I${r}=0,0,F${r}/I${r})` },                        // K: MTD Ach%
      11: { f: `IF(E${r}=0,0,F${r}/E${r})` },                        // L: SO vs Proj%
      12: { f: `IF(E${r}=0,0,1-L${r})` },                            // M: % SO Left
      13: { f: `IF(${DE}=0,0,L${r}/${DE})` },                        // N: SO%/Day
      14: { f: `IF(N${r}=0,0,M${r}/N${r})` },                        // O: Days Left
      15: { f: `I${r}-F${r}` },                                       // P: Diff KGs
      16: { f: `H${r}-J${r}` },                                       // Q: Diff Units
    };
  }

  // Build a placeholder row (17 cells)
  function placeholderRow(label1: string, label2: string, label3: string, label4: string): unknown[] {
    // E-Q will be overwritten with formulas later
    return [label1, label2, label3, label4, 0, 0, 0, null, null, 0, null, null, null, null, null, null, null];
  }

  // Fill summary formulas onto an existing aoa row
  function fillSummaryFormulas(rowArr: unknown[], excelRow: number, sumExpr: (col: string) => string) {
    rowArr[4] = { f: sumExpr('E') };   // Proj KGs
    rowArr[5] = { f: sumExpr('F') };   // SO KGs
    rowArr[6] = { f: sumExpr('G') };   // Proj Units
    rowArr[9] = { f: sumExpr('J') };   // SO Units
    const fCells = formulaCells(excelRow);
    rowArr[7]  = fCells[7];   // H: Exp SO Units
    rowArr[8]  = fCells[8];   // I: Exp KGs
    rowArr[10] = fCells[10];  // K: MTD Ach%
    rowArr[11] = fCells[11];  // L: SO vs Proj%
    rowArr[12] = fCells[12];  // M: % SO Left
    rowArr[13] = fCells[13];  // N: SO%/Day
    rowArr[14] = fCells[14];  // O: Days Left
    rowArr[15] = fCells[15];  // P: Diff KGs
    rowArr[16] = fCells[16];  // Q: Diff Units
  }

  sortedGroups.forEach(([gName, gRows]) => {
    const parentRowIdx = aoa.length;

    if (subGroupKey) {
      const subGroups: Record<string, Row[]> = {};
      gRows.forEach(r => {
        const sg = (r[subGroupKey] as string) || 'Unknown';
        if (!subGroups[sg]) subGroups[sg] = [];
        subGroups[sg].push(r);
      });
      const sortedSubs = Object.entries(subGroups).sort(([a], [b]) => a.localeCompare(b));

      aoa.push(placeholderRow(gName, '', '', `(${gRows.length} items)`));
      const parentExcelRow = aoa.length;
      const childStartRows: number[] = [];

      sortedSubs.forEach(([sgName, sgRows]) => {
        aoa.push(placeholderRow(gName, sgName, '', `(${sgRows.length})`));
        const subParentExcelRow = aoa.length;
        const subChildStart = aoa.length + 1;

        sgRows.forEach(r => {
          const exRow = aoa.length + 1;
          const fC = formulaCells(exRow);
          aoa.push([
            gName, sgName, r.itemCode, r.itemName,
            r.projKg, r.soKg, r.projUnits,
            fC[7], fC[8],          // H: Exp SO Units, I: Exp KGs
            r.soUnits,
            fC[10], fC[11], fC[12], fC[13], fC[14],  // K-O
            fC[15], fC[16],        // P-Q
          ]);
          outlineRows.push({ level: 2, row: aoa.length - 1 });
        });

        const subChildEnd = aoa.length;
        const sgr = aoa[subParentExcelRow - 1] as unknown[];
        fillSummaryFormulas(sgr, subParentExcelRow, col => `SUM(${col}${subChildStart}:${col}${subChildEnd})`);
        childStartRows.push(subParentExcelRow);
        outlineRows.push({ level: 1, row: subParentExcelRow - 1 });
      });

      const pr = aoa[parentRowIdx] as unknown[];
      fillSummaryFormulas(pr, parentExcelRow, col => childStartRows.map(r => `${col}${r}`).join('+'));

    } else {
      aoa.push(placeholderRow(gName, '', '', `(${gRows.length} items)`));
      const parentExcelRow = aoa.length;
      const childStart = aoa.length + 1;

      gRows.forEach(r => {
        const exRow = aoa.length + 1;
        const fC = formulaCells(exRow);
        aoa.push([
          gName, '', r.itemCode, r.itemName,
          r.projKg, r.soKg, r.projUnits,
          fC[7], fC[8],
          r.soUnits,
          fC[10], fC[11], fC[12], fC[13], fC[14],
          fC[15], fC[16],
        ]);
        outlineRows.push({ level: 1, row: aoa.length - 1 });
      });

      const childEnd = aoa.length;
      const pr = aoa[parentRowIdx] as unknown[];
      fillSummaryFormulas(pr, parentExcelRow, col => `SUM(${col}${childStart}:${col}${childEnd})`);
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Format percentage columns K, L, M, N (indices 10-13)
  const pctGroupCols = ['K', 'L', 'M', 'N'];
  for (let i = 1; i < aoa.length; i++) {
    const row = i + 1;
    pctGroupCols.forEach(c => { if (ws[`${c}${row}`]) ws[`${c}${row}`].z = '0.0%'; });
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

  ws['!autofilter'] = { ref: `A1:${colLetter(HC - 1)}1` };
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Summary sheet ───
// A=Name B=CustCount C=ProjKGs D=SOKGs E=ProjUnits F=ExpSOUnits
// G=ExpKGs H=SOUnits I=MTDAch% J=SOvsProj% K=%SOLeft L=DiffKGs M=DiffUnits
function writeSummarySheet(wb: XLSX.WorkBook, summaries: Summary[], sheetName: string) {
  const headers = [
    'Name', 'Cust Count', 'Proj KGs', 'SO KGs', 'Proj Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left',
    'Diff KGs', 'Diff Units'
  ];
  const aoa: unknown[][] = [headers];

  summaries.forEach(s => {
    const row = aoa.length + 1;
    aoa.push([
      s.name, s.count, s.projKg, s.soKg, s.projUnits,
      s.expKg, s.soUnits,
      { f: `IF(F${row}=0,0,D${row}/F${row})` },       // H: MTD Ach% = SOKg/ExpKg
      { f: `IF(C${row}=0,0,D${row}/C${row})` },       // I: SO vs Proj% = SOKg/ProjKg
      { f: `IF(C${row}=0,0,1-I${row})` },             // J: % SO Left
      { f: `F${row}-D${row}` },                        // K: Diff KGs = ExpKg - SOKg
      { f: `E${row}-G${row}` },                        // L: Diff Units
    ]);
  });

  // Totals row
  const totalRow = aoa.length + 1;
  const lastData = aoa.length;
  aoa.push([
    'TOTAL',
    { f: `SUM(B2:B${lastData})` },
    { f: `SUM(C2:C${lastData})` }, { f: `SUM(D2:D${lastData})` },
    { f: `SUM(E2:E${lastData})` },
    { f: `SUM(F2:F${lastData})` }, { f: `SUM(G2:G${lastData})` },
    { f: `IF(F${totalRow}=0,0,D${totalRow}/F${totalRow})` },
    { f: `IF(C${totalRow}=0,0,D${totalRow}/C${totalRow})` },
    { f: `IF(C${totalRow}=0,0,1-I${totalRow})` },
    { f: `F${totalRow}-D${totalRow}` },
    { f: `E${totalRow}-G${totalRow}` },
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const pctSumCols = ['H', 'I', 'J'];
  for (let i = 1; i <= summaries.length + 1; i++) {
    const row = i + 1;
    pctSumCols.forEach(c => { if (ws[`${c}${row}`]) ws[`${c}${row}`].z = '0.0%'; });
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
// A=Rank B=ItemCode C=ItemName D=CustGroup E=Customer F=Origin
// G=ProjKGs H=SOKGs I=ExpKGs J=MTDAch% K=SOvsProj% L=%SOLeft
// M=SO%/Day N=DaysLeft O=DiffKGs P=RunRateKGs
function writeLeaderboardSheet(wb: XLSX.WorkBook, rows: Row[], sheetName: string) {
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';
  const headers = [
    'Rank', 'Item Code', 'Item Name', 'Cust Group', 'Customer', 'Origin',
    'Proj KGs', 'SO KGs', 'Exp KGs',
    'MTD Ach%', 'SO vs Proj%', '% SO Left',
    'SO%/Day', 'Days Left', 'Diff KGs', 'RunRate KGs'
  ];
  const aoa: unknown[][] = [headers];

  rows.forEach((r, i) => {
    const row = i + 2;
    aoa.push([
      i + 1, r.itemCode, r.itemName, r.custGroup, r.customer, r.origin,
      r.projKg,                                                // G: Proj KGs
      r.soKg,                                                  // H: SO KGs
      { f: `IF(${DIM}=0,0,(G${row}/${DIM})*${DE})` },         // I: Exp KGs
      { f: `IF(I${row}=0,0,H${row}/I${row})` },               // J: MTD Ach%
      { f: `IF(G${row}=0,0,H${row}/G${row})` },               // K: SO vs Proj%
      { f: `IF(G${row}=0,0,1-K${row})` },                     // L: % SO Left
      { f: `IF(${DE}=0,0,K${row}/${DE})` },                   // M: SO%/Day
      { f: `IF(M${row}=0,0,L${row}/M${row})` },               // N: Days Left
      { f: `I${row}-H${row}` },                                // O: Diff KGs
      { f: `IF(${DE}=0,0,(H${row}/${DE})*${DIM})` },          // P: RunRate KGs
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const pctLBCols = ['J', 'K', 'L', 'M'];
  for (let i = 0; i < rows.length; i++) {
    const row = i + 2;
    pctLBCols.forEach(c => { if (ws[`${c}${row}`]) ws[`${c}${row}`].z = '0.0%'; });
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
// A=ItemName B=Customer C=ItemCode D=Origin
// E=ProjKGs F=SOKGs G=ProjUnits H=ExpSOUnits I=ExpKGs J=SOUnits
// K=MTDAch% L=SOvsProj% M=%SOLeft N=SO%/Day O=DaysLeft
// P=DiffKGs Q=DiffUnits
function writeCFASheet(wb: XLSX.WorkBook, data: DashData, cfaNames: Set<string>) {
  const cfaRows = data.rows.filter(r => cfaNames.has(r.itemName));
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';
  const headers = [
    'Item Name', 'Customer', 'Item Code', 'Origin',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units'
  ];
  const aoa: unknown[][] = [headers];
  cfaRows.forEach(r => {
    const row = aoa.length + 1;
    aoa.push([
      r.itemName, r.customer, r.itemCode, r.origin,
      r.projKg, r.soKg, r.projUnits,
      { f: `IF(${DIM}=0,0,(G${row}/${DIM})*${DE})` },   // H: Exp SO Units
      { f: `IF(${DIM}=0,0,(E${row}/${DIM})*${DE})` },   // I: Exp KGs
      r.soUnits,                                          // J
      { f: `IF(I${row}=0,0,F${row}/I${row})` },          // K: MTD Ach%
      { f: `IF(E${row}=0,0,F${row}/E${row})` },          // L: SO vs Proj%
      { f: `IF(E${row}=0,0,1-L${row})` },                // M: % SO Left
      { f: `IF(${DE}=0,0,L${row}/${DE})` },              // N: SO%/Day
      { f: `IF(N${row}=0,0,M${row}/N${row})` },          // O: Days Left
      { f: `I${row}-F${row}` },                           // P: Diff KGs
      { f: `H${row}-J${row}` },                           // Q: Diff Units
    ]);
  });

  // Totals
  const totalRow = aoa.length + 1;
  const last = aoa.length;
  aoa.push([
    'TOTAL', '', '', '',
    { f: `SUM(E2:E${last})` }, { f: `SUM(F2:F${last})` },
    { f: `SUM(G2:G${last})` }, null,
    null, { f: `SUM(J2:J${last})` },
    { f: `IF(I${totalRow}=0,0,F${totalRow}/I${totalRow})` },
    { f: `IF(E${totalRow}=0,0,F${totalRow}/E${totalRow})` },
    { f: `IF(E${totalRow}=0,0,1-L${totalRow})` },
    null, null,
    { f: `I${totalRow}-F${totalRow}` },
    { f: `H${totalRow}-J${totalRow}` },
  ]);
  // Fill SUM for Exp SO Units (H) and Exp KGs (I) on total row
  const totalArr = aoa[aoa.length - 1] as unknown[];
  totalArr[7] = { f: `IF(${DIM}=0,0,(G${totalRow}/${DIM})*${DE})` };
  totalArr[8] = { f: `IF(${DIM}=0,0,(E${totalRow}/${DIM})*${DE})` };

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const pctCFACols = ['K', 'L', 'M', 'N'];
  for (let i = 1; i < aoa.length; i++) {
    const row = i + 1;
    pctCFACols.forEach(c => { if (ws[`${c}${row}`]) ws[`${c}${row}`].z = '0.0%'; });
  }
  ws['!cols'] = [
    { wch: 45 }, { wch: 20 }, { wch: 18 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 13 },
    { wch: 12 }, { wch: 12 },
    { wch: 10 }, { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 12 }, { wch: 12 },
  ];
  ws['!autofilter'] = { ref: `A1:Q1` };
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
