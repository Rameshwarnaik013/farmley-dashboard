import * as XLSX from 'xlsx';
import { DashData, Row, Summary } from './helpers';

// Safe divide
function div(a: number, b: number): number { return b === 0 ? 0 : a / b; }
function rnd(v: number, d = 4): number { const m = Math.pow(10, d); return Math.round(v * m) / m; }

// Set a formula cell directly on the worksheet (bypasses aoa_to_sheet limitations)
function setF(ws: XLSX.WorkSheet, r: number, c: number, formula: string, cachedValue: number, fmt?: string) {
  const addr = XLSX.utils.encode_cell({ r, c });
  ws[addr] = { t: 'n', f: formula, v: cachedValue };
  if (fmt) ws[addr].z = fmt;
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
  // Col indices: H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15 Q=16 R=17 S=18 T=19 U=20 V=21
  const DIM = 'Config!$B$4';
  const DE = 'Config!$B$7';
  const PCT = '0.0%';

  // Build aoa with plain values first (formulas overlaid after)
  const aoa: (string | number)[][] = [headers];
  data.rows.forEach(r => {
    aoa.push([
      r.custGroup, r.customer, r.origin, r.itemCode, r.itemName, r.itemGroup, r.newMIS,
      r.projKg, r.soKg, r.projUnits,
      rnd(r.expUnits, 0),   // K placeholder
      rnd(r.expKg),         // L placeholder
      r.soUnits,
      rnd(r.achKg / 100),       // N placeholder (decimal for %)
      rnd(r.soVsProj / 100),    // O placeholder
      rnd(r.soLeftPct / 100),   // P placeholder
      rnd(r.soPctPerDay / 100), // Q placeholder
      rnd(r.daysToCover, 1),    // R placeholder
      rnd(r.diffKg),            // S placeholder
      rnd(r.diffUnits, 0),      // T placeholder
      rnd(r.runRateKg),         // U placeholder
      rnd(r.runRateUnits, 0),   // V placeholder
      r.isProj ? 'Yes' : 'No',
      r.isUnproj ? 'Yes' : 'No',
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Overlay formulas on each data row
  data.rows.forEach((r, i) => {
    const row = i + 2; // Excel 1-indexed, header is row 1
    const ri = i + 1;  // 0-indexed row in ws

    // K (col 10): Exp SO Units = IF(DIM=0,0,(J/DIM)*DE)
    setF(ws, ri, 10, `IF(${DIM}=0,0,(J${row}/${DIM})*${DE})`, rnd(r.expUnits, 0));
    // L (col 11): Exp KGs = IF(DIM=0,0,(H/DIM)*DE)
    setF(ws, ri, 11, `IF(${DIM}=0,0,(H${row}/${DIM})*${DE})`, rnd(r.expKg));
    // N (col 13): MTD Ach% = IF(L=0,0,I/L)
    setF(ws, ri, 13, `IF(L${row}=0,0,I${row}/L${row})`, rnd(div(r.soKg, r.expKg)), PCT);
    // O (col 14): SO vs Proj% = IF(H=0,0,I/H)
    setF(ws, ri, 14, `IF(H${row}=0,0,I${row}/H${row})`, rnd(div(r.soKg, r.projKg)), PCT);
    // P (col 15): % SO Left = IF(H=0,0,1-O)
    setF(ws, ri, 15, `IF(H${row}=0,0,1-O${row})`, rnd(r.projKg > 0 ? 1 - div(r.soKg, r.projKg) : 0), PCT);
    // Q (col 16): SO%/Day = IF(DE=0,0,O/DE)
    setF(ws, ri, 16, `IF(${DE}=0,0,O${row}/${DE})`, rnd(div(div(r.soKg, r.projKg), data.config.daysElapsed)), PCT);
    // R (col 17): Days Left = IF(Q=0,0,P/Q)
    setF(ws, ri, 17, `IF(Q${row}=0,0,P${row}/Q${row})`, rnd(r.daysToCover, 1));
    // S (col 18): Diff KGs = L-I
    setF(ws, ri, 18, `L${row}-I${row}`, rnd(r.expKg - r.soKg));
    // T (col 19): Diff Units = K-M
    setF(ws, ri, 19, `K${row}-M${row}`, rnd(r.expUnits - r.soUnits, 0));
    // U (col 20): RunRate KGs = IF(DE=0,0,(I/DE)*DIM)
    setF(ws, ri, 20, `IF(${DE}=0,0,(I${row}/${DE})*${DIM})`, rnd(r.runRateKg));
    // V (col 21): RunRate Units = IF(DE=0,0,(M/DE)*DIM)
    setF(ws, ri, 21, `IF(${DE}=0,0,(M${row}/${DE})*${DIM})`, rnd(r.runRateUnits, 0));
  });

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

// ─── Grouped pivot sheet with formulas ───
function writeGroupedSheet(
  wb: XLSX.WorkBook, data: DashData,
  sheetName: string, groupKey: keyof Row, subGroupKey?: keyof Row
) {
  const DIM = data.config.daysInMonth;
  const DE = data.config.daysElapsed;
  const DIM_REF = 'Config!$B$4';
  const DE_REF = 'Config!$B$7';
  const PCT = '0.0%';

  const headers = [
    'Group', 'Sub Group', 'Item Code', 'Item Name',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units'
  ];
  // Col: E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15 Q=16
  const HC = headers.length;

  function computeMetrics(projKg: number, soKg: number, projUnits: number, soUnits: number) {
    const expSoUnits = DIM === 0 ? 0 : rnd((projUnits / DIM) * DE, 0);
    const expKg = DIM === 0 ? 0 : rnd((projKg / DIM) * DE);
    return { expSoUnits, expKg, soUnits };
  }

  // Build aoa with plain values; track formula rows to overlay
  const aoa: (string | number)[][] = [headers];
  const outlineRows: { level: number; row: number }[] = [];

  // Track rows that need formulas: { wsRow, type, childRangeOrValues }
  interface FormulaInfo {
    wsRow: number; // 0-indexed worksheet row
    excelRow: number; // 1-indexed Excel row
    projKg: number; soKg: number; projUnits: number; soUnits: number;
    sumExpr?: (col: string) => string; // for aggregate rows
  }
  const formulaRows: FormulaInfo[] = [];

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

    if (subGroupKey) {
      const subGroups: Record<string, Row[]> = {};
      gRows.forEach(r => {
        const sg = (r[subGroupKey] as string) || 'Unknown';
        if (!subGroups[sg]) subGroups[sg] = [];
        subGroups[sg].push(r);
      });
      const sortedSubs = Object.entries(subGroups).sort(([a], [b]) => a.localeCompare(b));

      // Parent row placeholder
      const pProjKg = gRows.reduce((s, r) => s + r.projKg, 0);
      const pSoKg = gRows.reduce((s, r) => s + r.soKg, 0);
      const pProjUnits = gRows.reduce((s, r) => s + r.projUnits, 0);
      const pSoUnits = gRows.reduce((s, r) => s + r.soUnits, 0);
      const pm = computeMetrics(pProjKg, pSoKg, pProjUnits, pSoUnits);
      aoa.push([gName, '', '', `(${gRows.length} items)`, rnd(pProjKg), rnd(pSoKg), rnd(pProjUnits, 0), pm.expSoUnits, pm.expKg, rnd(pSoUnits, 0), 0, 0, 0, 0, 0, 0, 0]);
      const parentExcelRow = aoa.length; // 1-indexed
      const childStartRows: number[] = [];

      sortedSubs.forEach(([sgName, sgRows]) => {
        // Sub-group parent placeholder
        const sProjKg = sgRows.reduce((s, r) => s + r.projKg, 0);
        const sSoKg = sgRows.reduce((s, r) => s + r.soKg, 0);
        const sProjUnits = sgRows.reduce((s, r) => s + r.projUnits, 0);
        const sSoUnits = sgRows.reduce((s, r) => s + r.soUnits, 0);
        const sm = computeMetrics(sProjKg, sSoKg, sProjUnits, sSoUnits);
        aoa.push([gName, sgName, '', `(${sgRows.length})`, rnd(sProjKg), rnd(sSoKg), rnd(sProjUnits, 0), sm.expSoUnits, sm.expKg, rnd(sSoUnits, 0), 0, 0, 0, 0, 0, 0, 0]);
        const subParentExcelRow = aoa.length;
        const subChildStart = aoa.length + 1;
        outlineRows.push({ level: 1, row: aoa.length - 1 });

        sgRows.forEach(r => {
          const m = computeMetrics(r.projKg, r.soKg, r.projUnits, r.soUnits);
          aoa.push([gName, sgName, r.itemCode, r.itemName, r.projKg, r.soKg, r.projUnits, m.expSoUnits, m.expKg, r.soUnits, 0, 0, 0, 0, 0, 0, 0]);
          const curExcelRow = aoa.length;
          formulaRows.push({ wsRow: aoa.length - 1, excelRow: curExcelRow, projKg: r.projKg, soKg: r.soKg, projUnits: r.projUnits, soUnits: r.soUnits });
          outlineRows.push({ level: 2, row: aoa.length - 1 });
        });

        const subChildEnd = aoa.length;
        formulaRows.push({
          wsRow: subParentExcelRow - 1, excelRow: subParentExcelRow,
          projKg: sProjKg, soKg: sSoKg, projUnits: sProjUnits, soUnits: sSoUnits,
          sumExpr: col => `SUM(${col}${subChildStart}:${col}${subChildEnd})`,
        });
        childStartRows.push(subParentExcelRow);
      });

      formulaRows.push({
        wsRow: parentIdx, excelRow: parentExcelRow,
        projKg: pProjKg, soKg: pSoKg, projUnits: pProjUnits, soUnits: pSoUnits,
        sumExpr: col => childStartRows.map(r => `${col}${r}`).join('+'),
      });

    } else {
      const pProjKg = gRows.reduce((s, r) => s + r.projKg, 0);
      const pSoKg = gRows.reduce((s, r) => s + r.soKg, 0);
      const pProjUnits = gRows.reduce((s, r) => s + r.projUnits, 0);
      const pSoUnits = gRows.reduce((s, r) => s + r.soUnits, 0);
      const pm = computeMetrics(pProjKg, pSoKg, pProjUnits, pSoUnits);
      aoa.push([gName, '', '', `(${gRows.length} items)`, rnd(pProjKg), rnd(pSoKg), rnd(pProjUnits, 0), pm.expSoUnits, pm.expKg, rnd(pSoUnits, 0), 0, 0, 0, 0, 0, 0, 0]);
      const parentExcelRow = aoa.length;
      const childStart = aoa.length + 1;

      gRows.forEach(r => {
        const m = computeMetrics(r.projKg, r.soKg, r.projUnits, r.soUnits);
        aoa.push([gName, '', r.itemCode, r.itemName, r.projKg, r.soKg, r.projUnits, m.expSoUnits, m.expKg, r.soUnits, 0, 0, 0, 0, 0, 0, 0]);
        const curExcelRow = aoa.length;
        formulaRows.push({ wsRow: aoa.length - 1, excelRow: curExcelRow, projKg: r.projKg, soKg: r.soKg, projUnits: r.projUnits, soUnits: r.soUnits });
        outlineRows.push({ level: 1, row: aoa.length - 1 });
      });

      const childEnd = aoa.length;
      formulaRows.push({
        wsRow: parentIdx, excelRow: parentExcelRow,
        projKg: pProjKg, soKg: pSoKg, projUnits: pProjUnits, soUnits: pSoUnits,
        sumExpr: col => `SUM(${col}${childStart}:${col}${childEnd})`,
      });
    }
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Overlay formulas onto each tracked row
  formulaRows.forEach(info => {
    const { wsRow, excelRow: R, projKg, soKg, projUnits, soUnits } = info;
    const expKg = DIM === 0 ? 0 : (projKg / DIM) * DE;
    const soVsProj = div(soKg, projKg);
    const soLeft = projKg > 0 ? 1 - soVsProj : 0;
    const soPctDay = DE > 0 ? soVsProj / DE : 0;
    const daysLeft = soPctDay > 0 ? soLeft / soPctDay : 0;

    if (info.sumExpr) {
      // Aggregate row: E,F,G,J are SUMs
      setF(ws, wsRow, 4, info.sumExpr('E'), rnd(projKg));
      setF(ws, wsRow, 5, info.sumExpr('F'), rnd(soKg));
      setF(ws, wsRow, 6, info.sumExpr('G'), rnd(projUnits, 0));
      setF(ws, wsRow, 9, info.sumExpr('J'), rnd(soUnits, 0));
    }
    // H (7): Exp SO Units
    setF(ws, wsRow, 7, `IF(${DIM_REF}=0,0,(G${R}/${DIM_REF})*${DE_REF})`, rnd((projUnits / (DIM || 1)) * DE, 0));
    // I (8): Exp KGs
    setF(ws, wsRow, 8, `IF(${DIM_REF}=0,0,(E${R}/${DIM_REF})*${DE_REF})`, rnd(expKg));
    // K (10): MTD Ach%
    setF(ws, wsRow, 10, `IF(I${R}=0,0,F${R}/I${R})`, rnd(div(soKg, expKg)), PCT);
    // L (11): SO vs Proj%
    setF(ws, wsRow, 11, `IF(E${R}=0,0,F${R}/E${R})`, rnd(soVsProj), PCT);
    // M (12): % SO Left
    setF(ws, wsRow, 12, `IF(E${R}=0,0,1-L${R})`, rnd(soLeft), PCT);
    // N (13): SO%/Day
    setF(ws, wsRow, 13, `IF(${DE_REF}=0,0,L${R}/${DE_REF})`, rnd(soPctDay), PCT);
    // O (14): Days Left
    setF(ws, wsRow, 14, `IF(N${R}=0,0,M${R}/N${R})`, rnd(daysLeft, 1));
    // P (15): Diff KGs
    setF(ws, wsRow, 15, `I${R}-F${R}`, rnd(expKg - soKg));
    // Q (16): Diff Units
    setF(ws, wsRow, 16, `H${R}-J${R}`, rnd((projUnits / (DIM || 1)) * DE - soUnits, 0));
  });

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
  const DE = data.config.daysElapsed;
  const DE_REF = 'Config!$B$7';
  const PCT = '0.0%';
  const headers = [
    'Name', 'Cust Count', 'Proj KGs', 'SO KGs', 'Proj Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units'
  ];
  // A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13
  const aoa: (string | number)[][] = [headers];

  let tCount = 0, tProjKg = 0, tSoKg = 0, tProjUnits = 0, tExpKg = 0, tSoUnits = 0;

  summaries.forEach(s => {
    aoa.push([
      s.name, s.count, rnd(s.projKg), rnd(s.soKg), rnd(s.projUnits, 0),
      rnd(s.expKg), rnd(s.soUnits, 0),
      0, 0, 0, 0, 0, 0, 0
    ]);
    tCount += s.count; tProjKg += s.projKg; tSoKg += s.soKg;
    tProjUnits += s.projUnits; tExpKg += s.expKg; tSoUnits += s.soUnits;
  });

  // Totals row
  aoa.push([
    'TOTAL', tCount, rnd(tProjKg), rnd(tSoKg), rnd(tProjUnits, 0),
    rnd(tExpKg), rnd(tSoUnits, 0),
    0, 0, 0, 0, 0, 0, 0
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Overlay formulas for data rows + total row
  for (let i = 0; i < summaries.length + 1; i++) {
    const R = i + 2; // Excel row (1-indexed)
    const ri = i + 1; // ws row (0-indexed)
    const s = i < summaries.length ? summaries[i] : null;
    const projKg = s ? s.projKg : tProjKg;
    const soKg = s ? s.soKg : tSoKg;
    const expKg = s ? s.expKg : tExpKg;
    const soVsProj = div(soKg, projKg);
    const soLeft = projKg > 0 ? 1 - soVsProj : 0;
    const soPctDay = DE > 0 ? soVsProj / DE : 0;
    const daysLeft = soPctDay > 0 ? soLeft / soPctDay : 0;

    if (!s) {
      // Total row: also set SUM formulas for B-G
      const last = summaries.length + 1;
      setF(ws, ri, 1, `SUM(B2:B${last})`, tCount);
      setF(ws, ri, 2, `SUM(C2:C${last})`, rnd(tProjKg));
      setF(ws, ri, 3, `SUM(D2:D${last})`, rnd(tSoKg));
      setF(ws, ri, 4, `SUM(E2:E${last})`, rnd(tProjUnits, 0));
      setF(ws, ri, 5, `SUM(F2:F${last})`, rnd(tExpKg));
      setF(ws, ri, 6, `SUM(G2:G${last})`, rnd(tSoUnits, 0));
    }

    // H (7): MTD Ach% = SOKg/ExpKg
    setF(ws, ri, 7, `IF(F${R}=0,0,D${R}/F${R})`, rnd(div(soKg, expKg)), PCT);
    // I (8): SO vs Proj%
    setF(ws, ri, 8, `IF(C${R}=0,0,D${R}/C${R})`, rnd(soVsProj), PCT);
    // J (9): % SO Left
    setF(ws, ri, 9, `IF(C${R}=0,0,1-I${R})`, rnd(soLeft), PCT);
    // K (10): SO%/Day
    setF(ws, ri, 10, `IF(${DE_REF}=0,0,I${R}/${DE_REF})`, rnd(soPctDay), PCT);
    // L (11): Days Left
    setF(ws, ri, 11, `IF(K${R}=0,0,J${R}/K${R})`, rnd(daysLeft, 1));
    // M (12): Diff KGs
    setF(ws, ri, 12, `F${R}-D${R}`, rnd(expKg - soKg));
    // N (13): Diff Units
    setF(ws, ri, 13, `E${R}-G${R}`, rnd((s ? s.projUnits : tProjUnits) - (s ? s.soUnits : tSoUnits), 0));
  }

  ws['!cols'] = [
    { wch: 28 }, { wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
    { wch: 14 }, { wch: 14 },
    { wch: 10 }, { wch: 11 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 14 }, { wch: 14 },
  ];
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
}

// ─── Leaderboard sheet ───
function writeLeaderboardSheet(wb: XLSX.WorkBook, rows: Row[], sheetName: string, data: DashData) {
  const DIM_REF = 'Config!$B$4';
  const DE_REF = 'Config!$B$7';
  const DE = data.config.daysElapsed;
  const DIM = data.config.daysInMonth;
  const PCT = '0.0%';
  const headers = [
    'Rank', 'Item Code', 'Item Name', 'Cust Group', 'Customer', 'Origin',
    'Proj KGs', 'SO KGs', 'Exp KGs',
    'MTD Ach%', 'SO vs Proj%', '% SO Left',
    'SO%/Day', 'Days Left', 'Diff KGs', 'RunRate KGs'
  ];
  // G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15
  const aoa: (string | number)[][] = [headers];

  rows.forEach((r, i) => {
    aoa.push([
      i + 1, r.itemCode, r.itemName, r.custGroup, r.customer, r.origin,
      r.projKg, r.soKg, rnd(r.expKg),
      0, 0, 0, 0, 0, 0, 0
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  rows.forEach((r, i) => {
    const R = i + 2;
    const ri = i + 1;
    const expKg = DIM === 0 ? 0 : (r.projKg / DIM) * DE;
    const soVsProj = div(r.soKg, r.projKg);
    const soLeft = r.projKg > 0 ? 1 - soVsProj : 0;
    const soPctDay = DE > 0 ? soVsProj / DE : 0;
    const daysLeft = soPctDay > 0 ? soLeft / soPctDay : 0;

    // I (8): Exp KGs
    setF(ws, ri, 8, `IF(${DIM_REF}=0,0,(G${R}/${DIM_REF})*${DE_REF})`, rnd(expKg));
    // J (9): MTD Ach%
    setF(ws, ri, 9, `IF(I${R}=0,0,H${R}/I${R})`, rnd(div(r.soKg, expKg)), PCT);
    // K (10): SO vs Proj%
    setF(ws, ri, 10, `IF(G${R}=0,0,H${R}/G${R})`, rnd(soVsProj), PCT);
    // L (11): % SO Left
    setF(ws, ri, 11, `IF(G${R}=0,0,1-K${R})`, rnd(soLeft), PCT);
    // M (12): SO%/Day
    setF(ws, ri, 12, `IF(${DE_REF}=0,0,K${R}/${DE_REF})`, rnd(soPctDay), PCT);
    // N (13): Days Left
    setF(ws, ri, 13, `IF(M${R}=0,0,L${R}/M${R})`, rnd(daysLeft, 1));
    // O (14): Diff KGs
    setF(ws, ri, 14, `I${R}-H${R}`, rnd(expKg - r.soKg));
    // P (15): RunRate KGs
    setF(ws, ri, 15, `IF(${DE_REF}=0,0,(H${R}/${DE_REF})*${DIM_REF})`, rnd(r.runRateKg));
  });

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
  const DIM_REF = 'Config!$B$4';
  const DE_REF = 'Config!$B$7';
  const PCT = '0.0%';
  const headers = [
    'Item Name', 'Customer', 'Item Code', 'Origin',
    'Proj KGs', 'SO KGs', 'Proj Units', 'Exp SO Units',
    'Exp KGs', 'SO Units',
    'MTD Ach%', 'SO vs Proj%', '% SO Left', 'SO%/Day', 'Days Left',
    'Diff KGs', 'Diff Units'
  ];
  // E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15 Q=16
  const aoa: (string | number)[][] = [headers];

  let tProjKg = 0, tSoKg = 0, tProjUnits = 0, tSoUnits = 0;

  cfaRows.forEach(r => {
    const expSoUnits = DIM === 0 ? 0 : rnd((r.projUnits / DIM) * DE, 0);
    const expKg = DIM === 0 ? 0 : rnd((r.projKg / DIM) * DE);
    aoa.push([
      r.itemName, r.customer, r.itemCode, r.origin,
      r.projKg, r.soKg, r.projUnits, expSoUnits, expKg, r.soUnits,
      0, 0, 0, 0, 0, 0, 0
    ]);
    tProjKg += r.projKg; tSoKg += r.soKg; tProjUnits += r.projUnits; tSoUnits += r.soUnits;
  });

  // Totals row
  const tExpSoUnits = DIM === 0 ? 0 : rnd((tProjUnits / DIM) * DE, 0);
  const tExpKg = DIM === 0 ? 0 : rnd((tProjKg / DIM) * DE);
  aoa.push([
    'TOTAL', '', '', '',
    rnd(tProjKg), rnd(tSoKg), rnd(tProjUnits, 0), tExpSoUnits, tExpKg, rnd(tSoUnits, 0),
    0, 0, 0, 0, 0, 0, 0
  ]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Overlay formulas for each CFA row + total
  for (let i = 0; i < cfaRows.length + 1; i++) {
    const R = i + 2;
    const ri = i + 1;
    const isTotal = i === cfaRows.length;
    const r = isTotal ? null : cfaRows[i];
    const projKg = isTotal ? tProjKg : r!.projKg;
    const soKg = isTotal ? tSoKg : r!.soKg;
    const projUnits = isTotal ? tProjUnits : r!.projUnits;
    const soUnits = isTotal ? tSoUnits : r!.soUnits;
    const expKg = DIM === 0 ? 0 : (projKg / DIM) * DE;
    const soVsProj = div(soKg, projKg);
    const soLeft = projKg > 0 ? 1 - soVsProj : 0;
    const soPctDay = DE > 0 ? soVsProj / DE : 0;
    const daysLeft = soPctDay > 0 ? soLeft / soPctDay : 0;

    if (isTotal) {
      const last = cfaRows.length + 1;
      setF(ws, ri, 4, `SUM(E2:E${last})`, rnd(tProjKg));
      setF(ws, ri, 5, `SUM(F2:F${last})`, rnd(tSoKg));
      setF(ws, ri, 6, `SUM(G2:G${last})`, rnd(tProjUnits, 0));
      setF(ws, ri, 9, `SUM(J2:J${last})`, rnd(tSoUnits, 0));
    }

    // H (7): Exp SO Units
    setF(ws, ri, 7, `IF(${DIM_REF}=0,0,(G${R}/${DIM_REF})*${DE_REF})`, rnd((projUnits / (DIM || 1)) * DE, 0));
    // I (8): Exp KGs
    setF(ws, ri, 8, `IF(${DIM_REF}=0,0,(E${R}/${DIM_REF})*${DE_REF})`, rnd(expKg));
    // K (10): MTD Ach%
    setF(ws, ri, 10, `IF(I${R}=0,0,F${R}/I${R})`, rnd(div(soKg, expKg)), PCT);
    // L (11): SO vs Proj%
    setF(ws, ri, 11, `IF(E${R}=0,0,F${R}/E${R})`, rnd(soVsProj), PCT);
    // M (12): % SO Left
    setF(ws, ri, 12, `IF(E${R}=0,0,1-L${R})`, rnd(soLeft), PCT);
    // N (13): SO%/Day
    setF(ws, ri, 13, `IF(${DE_REF}=0,0,L${R}/${DE_REF})`, rnd(soPctDay), PCT);
    // O (14): Days Left
    setF(ws, ri, 14, `IF(N${R}=0,0,M${R}/N${R})`, rnd(daysLeft, 1));
    // P (15): Diff KGs
    setF(ws, ri, 15, `I${R}-F${R}`, rnd(expKg - soKg));
    // Q (16): Diff Units
    setF(ws, ri, 16, `H${R}-J${R}`, rnd((projUnits / (DIM || 1)) * DE - soUnits, 0));
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
  writeLeaderboardSheet(wb, data.top20Kg, 'Top 20 KGs', data);
  writeLeaderboardSheet(wb, data.top20Units, 'Top 20 Units', data);
  writeLeaderboardSheet(wb, data.bot20Kg, 'Bottom 20 KGs', data);
  writeLeaderboardSheet(wb, data.bot20Units, 'Bottom 20 Units', data);
  writeZeroSOSheet(wb, data);
  writeUnprojSheet(wb, data);

  const filename = `RunRate_Dashboard_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, filename);
}
