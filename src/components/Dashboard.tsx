'use client';
import { useState, useMemo, useCallback, useRef, Fragment } from 'react';
import { DashData, Row, fmt } from '@/lib/helpers';
import { exportDashboardExcel } from '@/lib/exportExcel';
import { aggGroup } from '@/lib/processData';
import StatsCards from './StatsCards';
import FilterBar from './FilterBar';
import DataTable from './DataTable';
import GroupedView from './GroupedView';
import Leaderboard from './Leaderboard';
import SummaryChart from './SummaryChart';

type Tab = 'ach-above-100' | 'instructions' | 'dashboard' | 'cfa-items' | 'new-mis' | 'cust-group' | 'origin' | 'item-name' | 'top-kg' | 'top-unit' | 'bot-kg' | 'bot-unit' | 'summary-ig' | 'summary-origin' | 'zero-so' | 'unproj' | 'data-audit';

const CFA_ITEM_NAMES = new Set([
  'Apple Pie Date Bite Farmley sachet 20 g - Single serve',
  'Apple Pie Date Bite Farmley Tin Jar 200g',
  'Apple Pie Date Bites Farmley Monocarton 240 g - Pack of 12 (20 g Each)',
  'Assorted Pack Farmley 60g Dark Choco Orange Date Bite 20g, Classic Date Bite 20g, Apple Pie Date Bite 20g',
  'Berry Mix Farmley Standee Pouch 200 g',
  'Classic Date Bites Farmley Monocarton 240 g - Pack of 12 (20 g Each)',
  'Classic Date Bites Farmley Sachet 20 g - Single Serve',
  'Classic Date Bites Farmley Tin Jar 200 g',
  'Coffee Rush Date Bite Farmley Sachet 20g- Single Serve',
  'Coffee Rush Date Bite Farmley Tin Jar 200g',
  'Dark Choco Orange Date Bites Farmley Sachet 20 g - Single Serve',
  'Dark Choco-Orange Date Bite Farmley Tin Jar 200g',
  'Dark Choco-Orange Date Bites Monocarton 240 g - Pack of 12 (20 g Each)',
  'Fruit and Nut Mix Farmley Pillow Pouch 28g- Single serve',
  'Mexican Peri Peri Snack Mix Farmley Composite Jar 325 g',
  'Mexican Peri Peri Snack Mix Farmley 21g - Pack of 10 (Ladi)',
  'Mexican Peri Peri Snack Mix Farmley Pillow Pouch 21g (In SRP 8 Pcs Each)',
  'Mexican Peri-Peri Snack Mix Farmley Pillow Pouch 35 g- Single Serve',
  'Prasadam Makhana Center Seal Pouch 100 g- NEW',
  'Prasadam Makhana Center Seal Pouch 200 g- NEW',
  'Prasadam Makhana Farmley Center Seal Pouch 60g-New',
  'Premium Panchmeva Farmley Jar 405 g',
  'Premium Panchmeva Farmley Jar 425 g- Institutional',
  'Premium Panchmeva Farmley Pillow Pouch 21 g - Pack of 10 (Ladi)',
  'Premium Panchmeva Farmley Pillow Pouch 30g- single serve',
  'Premium Panchmeva Farmley Pillow Pouch 21g (In SRP 8 Pcs Each)',
  'Roasted & Flavored Makhana - Achaari Farmley Composite Jar 77 g',
  'Roasted & Flavored Makhana - Achaari Farmley Pillow Pouch 20 g',
  'Roasted & Flavored Makhana - Black Pepper Farmley Composite jar 77 g',
  'Roasted & Flavored Makhana - Cream & Onion Farmley Pillow Pouch 20 g',
  'Roasted & Flavored Makhana - Mint Farmley Composite Jar 77 g',
  'Roasted & Flavored Makhana - Peri Peri Farmley Composite Jar 77 g',
  'Roasted & Flavored Makhana - Peri Peri Farmley Pillow Pouch 14 g-Ladi',
  'Roasted & Flavored Makhana - Peri Peri Farmley Pillow Pouch 20 g',
  'Roasted & Flavored Makhana - Tangy Tomato Farmley Composite Jar 77 g',
  'Roasted & Flavored Makhana - Tangy Tomato Farmley Pillow Pouch 20 g',
  'Roasted & Flavored Makhana -Cheddar Cheese Farmley Composite Jar 77 g',
  'Roasted & Flavored Makhana -Cream and Onion Farmley Composite Jar 77 g',
  'Roasted & Flavoured Achaari Makhana Farmley Jar 55 g',
  'Roasted & Flavoured Cheddar Cheese Makhana Farmley Jar 55 g',
  'Roasted & Flavoured Cream & Onion Makhana Farmley Jar 55 g',
  'Roasted & Flavoured Makhana - Achaari Farmley Pillow Pouch 14g-Ladi',
  'Roasted & Flavoured Makhana - Cream & Onion Farmley Pillow Pouch 14g-Ladi',
  'Roasted & Flavoured Makhana - Tangy Tomato Farmley Pillow Pouch 14g-Ladi',
  'Roasted & Flavoured Minty Pudina Makhana Farmley Jar 55 g',
  'Roasted & Flavoured Peri Peri Makhana Farmley Jar 55 g',
  'Roasted & Flavoured Salt & Pepper Makhana Farmley Jar 55 g',
  'Roasted & Flavoured Tangy Tomato Makhana Farmley Jar 55 g',
  'Roasted & Salted Makhana Farmley Composite Jar 77 g',
  'Roasted & Salted Makhana Farmley Jar 55 g',
  'Roasted & Salted Makhana Farmley Pillow Pouch 20 g',
  'Seed Mix Farmley Standee Pouch 200 g',
  'Smokey BBQ Party Mix Pillow Pouch Farmley 28 g- Single serve',
  'Sweet and Salty Mix Farmley Pillow Pouch 23g (In SRP 8 Pcs Each)',
  'Trail Mix Farmley Composite Jar 325 g',
  'Trail Mix Farmley Standee Pouch 200 g',
]);

const TABS: { id: Tab; label: string }[] = [
  { id: 'ach-above-100', label: 'MTD Ach% > 100%' },
  { id: 'item-name', label: 'By Item Name' },
  { id: 'instructions', label: 'Instructions' },
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'cfa-items', label: 'CFA Items Run Rate' },
  { id: 'new-mis', label: 'By New MIS' },
  { id: 'cust-group', label: 'By Customer Group' },
  { id: 'origin', label: 'By Origin' },
  { id: 'top-kg', label: 'Top 20 (KGs)' },
  { id: 'top-unit', label: 'Top 20 (Units)' },
  { id: 'bot-kg', label: 'Bottom 20 (KGs)' },
  { id: 'bot-unit', label: 'Bottom 20 (Units)' },
  { id: 'summary-ig', label: 'Summary: Item Group' },
  { id: 'summary-origin', label: 'Summary: Origin' },
  { id: 'zero-so', label: 'Zero SO' },
  { id: 'unproj', label: 'Unprojected SO' },
  { id: 'data-audit', label: 'Data Audit' },
];

interface DashboardProps {
  data: DashData;
  onReUpload: () => void;
}

function InstructionsPanel({ cfg }: { cfg: DashData['config'] }) {
  const sections = [
    { id: 'flow', title: '1. Data Processing Flow' },
    { id: 'inputs', title: '2. Input Files' },
    { id: 'cleaning', title: '3. SO Data Cleaning Rules (Applied on Input 2)' },
    { id: 'remap', title: '4. Customer Group / Customer Remapping' },
    { id: 'join', title: '5. Join Logic (Projection vs Sales Order)' },
    { id: 'formulas', title: '6. Column Formulas & Definitions' },
    { id: 'colors', title: '7. Color Coding (Procurement Focus)' },
    { id: 'filters', title: '8. Default Filters' },
    { id: 'tabs', title: '9. Tab Descriptions' },
    { id: 'excel', title: '10. Excel Export' },
  ];
  const [openSections, setOpenSections] = useState<Set<string>>(new Set(sections.map(s => s.id)));
  const toggle = (id: string) => setOpenSections(prev => {
    const next = new Set(prev);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  const S = ({ id, title, children }: { id: string; title: string; children: React.ReactNode }) => (
    <div className="border rounded-lg overflow-hidden">
      <button onClick={() => toggle(id)} className="w-full flex items-center justify-between px-4 py-3 bg-gray-50 hover:bg-gray-100 transition-colors text-left">
        <span className="text-sm font-bold text-brand-900">{title}</span>
        <span className="text-gray-400 text-xs">{openSections.has(id) ? '▼' : '▶'}</span>
      </button>
      {openSections.has(id) && <div className="px-4 py-3 text-xs leading-relaxed text-gray-700 space-y-2">{children}</div>}
    </div>
  );

  const tbl = "w-full text-xs border-collapse my-2";
  const th = "text-left px-3 py-1.5 bg-gray-100 border font-semibold";
  const td = "px-3 py-1.5 border";
  const code = "bg-gray-100 px-1.5 py-0.5 rounded font-mono text-[11px]";

  return (
    <div className="space-y-3 max-w-5xl">
      <div className="bg-white rounded-lg shadow-sm p-5 mb-4">
        <h2 className="text-base font-bold text-brand-900 mb-1">Dashboard Instructions & Methodology</h2>
        <p className="text-xs text-gray-500">Complete reference of all approaches, conditions, formulas, and logic used in this Run Rate Dashboard.</p>
        <p className="text-xs text-gray-400 mt-1">Current config: Days in Month = <b>{cfg.daysInMonth}</b> | Days Elapsed = <b>{cfg.daysElapsed}</b> | SO Date Range: {cfg.dateFrom} to {cfg.dateTo}</p>
      </div>

      <S id="flow" title="1. Data Processing Flow">
        <div className="bg-white rounded p-3 border text-[11px] font-mono leading-6">
          <p>Upload Projection Excel (Input 1) + Sales Order Excel (Input 2) + Days in Month</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>Apply 9 Cleaning Rules on SO data (remove invalid rows)</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>Apply Customer Group / Customer Remapping (both files)</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>Detect Date Range: 1st of month to max(Sales Order Date) → Days Elapsed = day of max date</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>Build SO Lookup: GROUP BY (Customer Group, Customer, Origin, Item Code) → SUM(Stock Qty), SUM(Qty)</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>OUTER JOIN Projection rows with SO Lookup on (Customer Group, Customer, Origin, Item Code)</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>Compute derived columns: Expected, MTD Ach%, SO vs Proj%, Days Left, Run Rate, etc.</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>Add Unprojected SO items (SO entries with no matching Projection row)</p>
          <p className="text-gray-400 my-1">{'  ↓'}</p>
          <p>Render Dashboard with filters, grouped views, summaries, leaderboards</p>
        </div>
        <p className="text-gray-500 mt-2">All processing happens entirely in the browser. No data is sent to any server.</p>
      </S>

      <S id="inputs" title="2. Input Files">
        <table className={tbl}>
          <thead><tr><th className={th}>Input</th><th className={th}>Description</th><th className={th}>Key Columns Used</th></tr></thead>
          <tbody>
            <tr><td className={td}><b>Input 1: Projection</b></td><td className={td}>Monthly projection/target data per SKU per customer</td><td className={td}>Customer Group, Customer, Origin, Item Code, Item Name, Total KGs, Projection Units, Item Group, Month, Year</td></tr>
            <tr><td className={td}><b>Input 2: Sales Order</b></td><td className={td}>Actual sales order line items for the month</td><td className={td}>Customer Group, Customer, Origin, Item Code, Item Name, Stock Qty (KGs), Qty (Units), Sales Order Date, Workflow State, Sales Order Status, Sales Order Created By, Purpose, Item Type, Returned Qty, NEW MIS ITEM GROUP</td></tr>
            <tr><td className={td}><b>Days in Month</b></td><td className={td}>Working days in the current month (user input)</td><td className={td}>Used as denominator for daily demand and expected calculations</td></tr>
          </tbody>
        </table>
      </S>

      <S id="cleaning" title="3. SO Data Cleaning Rules (Applied on Input 2)">
        <p className="mb-2">These 9 rules are applied <b>sequentially</b> on the raw Sales Order data. A row is removed on the <b>first matching rule</b> (no double-counting).</p>
        <table className={tbl}>
          <thead><tr><th className={th}>#</th><th className={th}>Rule</th><th className={th}>Condition</th></tr></thead>
          <tbody>
            <tr><td className={td}>1</td><td className={td}>Remove Administrator rows</td><td className={td}><span className={code}>Sales Order Created By == &quot;Administrator&quot;</span></td></tr>
            <tr><td className={td}>2</td><td className={td}>Remove invalid Internal Transfer (WF State)</td><td className={td}><span className={code}>Workflow State == &quot;Internal Transfer&quot;</span> AND creator NOT in [ashish.k@farmley.com, abhishek.ku@farmley.com]</td></tr>
            <tr><td className={td}>3</td><td className={td}>Remove On Hold / Rejected / Cancelled (WF State)</td><td className={td}><span className={code}>Workflow State IN (&quot;On Hold&quot;, &quot;Rejected&quot;, &quot;Cancelled&quot;)</span></td></tr>
            <tr><td className={td}>4</td><td className={td}>Remove Documentation purpose</td><td className={td}><span className={code}>Purpose == &quot;Documentation&quot;</span></td></tr>
            <tr><td className={td}>5</td><td className={td}>Remove Cancelled / On Hold / Rejected (SO Status)</td><td className={td}><span className={code}>Sales Order Status IN (&quot;Cancelled&quot;, &quot;On Hold&quot;, &quot;Rejected&quot;)</span></td></tr>
            <tr><td className={td}>6</td><td className={td}>Remove invalid Internal Transfer (SO Status)</td><td className={td}><span className={code}>Sales Order Status == &quot;Internal Transfer&quot;</span> AND creator NOT in valid list</td></tr>
            <tr><td className={td}>7</td><td className={td}>Remove Sample Orders</td><td className={td}><span className={code}>Customer STARTS WITH &quot;Sample Order&quot;</span></td></tr>
            <tr><td className={td}>8</td><td className={td}>Remove Bulk items</td><td className={td}><span className={code}>Item Type == &quot;Bulk&quot;</span></td></tr>
            <tr><td className={td}>9</td><td className={td}>Remove high-return rows</td><td className={td}><span className={code}>Returned Qty &gt; 1</span></td></tr>
          </tbody>
        </table>
        <p className="text-gray-500 mt-1">See the <b>Data Audit</b> tab for per-rule removal stats on your current data.</p>
      </S>

      <S id="remap" title="4. Customer Group / Customer Remapping">
        <p className="mb-2">Applied to <b>both Projection and SO data</b> before the join, ensuring consistent keys.</p>
        <table className={tbl}>
          <thead><tr><th className={th}>#</th><th className={th}>Condition</th><th className={th}>Action</th></tr></thead>
          <tbody>
            <tr><td className={td}>1</td><td className={td}>Customer Group = &quot;Category A&quot; AND Customer = &quot;Jhabak Marketing&quot;</td><td className={td}>Change Customer Group to <b>&quot;Modern Trade&quot;</b></td></tr>
            <tr><td className={td}>2</td><td className={td}>Customer Group = &quot;Category A&quot; AND Customer = any other</td><td className={td}>Change Customer to <b>&quot;GT Retail&quot;</b></td></tr>
            <tr><td className={td}>3</td><td className={td}>Customer Group = &quot;CPC KPKB&quot;</td><td className={td}>Change Customer to <b>&quot;CPC KPKB&quot;</b></td></tr>
            <tr><td className={td}>4</td><td className={td}>Customer Group = &quot;Quick Commerce&quot; AND Customer IN (Moksh Enterprises, PJTJ Technologies, Cloudkart Ventures, Jupiter Kart, Cloudstore Retail)</td><td className={td}>Change Customer to <b>&quot;Instamart&quot;</b></td></tr>
          </tbody>
        </table>
      </S>

      <S id="join" title="5. Join Logic (Projection vs Sales Order)">
        <p><b>Join Key:</b> <span className={code}>(Customer Group, Customer, Origin, Item Code)</span></p>
        <p className="mt-2"><b>SO Aggregation:</b> Before joining, SO data is grouped by the join key with <span className={code}>SUM(Stock Qty)</span> for SO KGs and <span className={code}>SUM(Qty)</span> for SO Units.</p>
        <p className="mt-2"><b>OUTER JOIN:</b></p>
        <ul className="list-disc pl-5 space-y-1 mt-1">
          <li><b>Projected rows:</b> Each Projection row is enriched with SO KGs/Units from the matching SO group. If no SO match exists, SO KGs = 0.</li>
          <li><b>Unprojected rows:</b> SO groups with no matching Projection row are added as separate rows (Proj KGs = 0, marked as &quot;Unprojected&quot;).</li>
        </ul>
        <p className="mt-2"><b>Item Group / New MIS:</b> <span className={code}>NEW MIS ITEM GROUP</span> in the SO file and <span className={code}>Item Group</span> in the Projection file are the same taxonomy, so they are treated as one dimension. Projected rows take Item Group from the Projection file; SO-only rows take it from <span className={code}>NEW MIS ITEM GROUP</span>. The SO file&apos;s own <span className={code}>Item Group</span> column (a coarser list — &quot;Dry Fruit Mix&quot;, &quot;Desserts&quot;, &quot;Makhana&quot;) is <b>not</b> used, so a single vocabulary applies across projected and unprojected rows.</p>
        <p className="mt-2"><b>Days Elapsed:</b> Automatically derived as the day-of-month of <span className={code}>MAX(Sales Order Date)</span>. Current value = <b>{cfg.daysElapsed}</b></p>
      </S>

      <S id="formulas" title="6. Column Formulas & Definitions">
        <table className={tbl}>
          <thead><tr><th className={th}>Column</th><th className={th}>Formula / Definition</th></tr></thead>
          <tbody>
            <tr><td className={td}><b>Proj KGs</b></td><td className={td}>Total KGs from Projection file (monthly target)</td></tr>
            <tr><td className={td}><b>SO KGs</b></td><td className={td}>SUM(Stock Qty) from SO data grouped by join key</td></tr>
            <tr><td className={td}><b>Proj Units</b></td><td className={td}>Projection Units from Projection file</td></tr>
            <tr><td className={td}><b>Exp SO Units</b></td><td className={td}><span className={code}>(Proj Units / Days in Month) x Days Elapsed</span></td></tr>
            <tr><td className={td}><b>Exp KGs</b></td><td className={td}><span className={code}>(Proj KGs / Days in Month) x Days Elapsed</span></td></tr>
            <tr><td className={td}><b>SO Units</b></td><td className={td}>SUM(Qty) from SO data grouped by join key</td></tr>
            <tr><td className={td}><b>MTD Ach%</b></td><td className={td}><span className={code}>SO KGs / Exp KGs x 100</span> — Month-to-date achievement (expected projection vs actual SO)</td></tr>
            <tr><td className={td}><b>SO vs Proj%</b></td><td className={td}><span className={code}>SO KGs / Proj KGs x 100</span> — How much of full-month projection is already covered</td></tr>
            <tr><td className={td}><b>% SO Left</b></td><td className={td}><span className={code}>100 - SO vs Proj%</span> — Remaining percentage of projection to be fulfilled</td></tr>
            <tr><td className={td}><b>SO%/Day</b></td><td className={td}><span className={code}>SO vs Proj% / Days Elapsed</span> — Average projection coverage per day</td></tr>
            <tr><td className={td}><b>Days Left</b></td><td className={td}><span className={code}>% SO Left / SO%/Day</span> — Estimated days needed to cover remaining projection at current pace</td></tr>
            <tr><td className={td}><b>Diff KGs</b></td><td className={td}><span className={code}>Exp KGs - SO KGs</span> — Positive = under-ordered, Negative = over-ordered</td></tr>
            <tr><td className={td}><b>Diff Units</b></td><td className={td}><span className={code}>Exp SO Units - SO Units</span></td></tr>
            <tr><td className={td}><b>RunRate Units</b></td><td className={td}><span className={code}>(SO Units / Days Elapsed) x Days in Month</span></td></tr>
          </tbody>
        </table>
      </S>

      <S id="colors" title="7. Color Coding (Procurement Focus)">
        <p className="mb-2">Colors indicate procurement urgency. <b>RED = SO exceeding projection</b> (need to arrange more stock).</p>
        <table className={tbl}>
          <thead><tr><th className={th}>Color</th><th className={th}>MTD Ach% Range</th><th className={th}>Meaning</th></tr></thead>
          <tbody>
            <tr><td className={td}><span className="inline-block w-4 h-4 rounded bg-[#c00000] align-middle mr-2"></span> Dark Red</td><td className={td}>&gt; 120%</td><td className={td}>Urgent — SO far exceeds projection pace, immediate procurement action needed</td></tr>
            <tr><td className={td}><span className="inline-block w-4 h-4 rounded bg-[#ff0000] align-middle mr-2"></span> Red</td><td className={td}>100% - 120%</td><td className={td}>Exceeding — SO ahead of projection, monitor closely</td></tr>
            <tr><td className={td}><span className="inline-block w-4 h-4 rounded bg-[#ffc000] align-middle mr-2"></span> Orange</td><td className={td}>80% - 100%</td><td className={td}>Watch — On track but nearing projection</td></tr>
            <tr><td className={td}><span className="inline-block w-4 h-4 rounded bg-[#92d050] align-middle mr-2"></span> Light Green</td><td className={td}>50% - 80%</td><td className={td}>Comfortable — SO below expected pace</td></tr>
            <tr><td className={td}><span className="inline-block w-4 h-4 rounded bg-[#00b050] align-middle mr-2"></span> Dark Green</td><td className={td}>&lt; 50%</td><td className={td}>Low — SO significantly below expected pace</td></tr>
            <tr><td className={td}><span className="inline-block w-4 h-4 rounded bg-[#d9d9d9] align-middle mr-2"></span> Grey</td><td className={td}>0% / No SO</td><td className={td}>No sales orders received yet</td></tr>
          </tbody>
        </table>
      </S>

      <S id="filters" title="8. Default Filters">
        <ul className="list-disc pl-5 space-y-1">
          <li><b>Customer Group:</b> All selected EXCEPT &quot;Inter Company&quot; (unchecked by default). Toggle it on to include inter-company transfers.</li>
          <li><b>Origin:</b> All selected by default.</li>
          <li><b>Item Group:</b> All selected by default.</li>
          <li>Filters apply to all tabs except Data Audit. Summary charts and leaderboards are dynamically recomputed based on active filters.</li>
          <li>Search (300ms debounce) filters by Item Name, Item Code, Customer, Customer Group, and New MIS.</li>
        </ul>
      </S>

      <S id="tabs" title="9. Tab Descriptions">
        <table className={tbl}>
          <thead><tr><th className={th}>Tab</th><th className={th}>Description</th></tr></thead>
          <tbody>
            <tr><td className={td}><b>Dashboard</b></td><td className={td}>Flat table of all rows with all columns. Sortable by any column. Paginated (50 rows/page).</td></tr>
            <tr><td className={td}><b>CFA Items Run Rate</b></td><td className={td}>Grouped by Item Name → Customer sub-toggle. Filtered to 56 specific CFA item names only.</td></tr>
            <tr><td className={td}><b>By New MIS</b></td><td className={td}>Grouped by NEW MIS ITEM GROUP. Expandable to see individual items.</td></tr>
            <tr><td className={td}><b>By Customer Group</b></td><td className={td}>Grouped by Customer Group → Customer sub-toggle. Shows aggregated metrics per group.</td></tr>
            <tr><td className={td}><b>By Origin</b></td><td className={td}>Grouped by Origin. Expandable to individual items.</td></tr>
            <tr><td className={td}><b>By Item Name</b></td><td className={td}>Grouped by Item Name → Customer sub-toggle. Shows per-item SO split by customer.</td></tr>
            <tr><td className={td}><b>Top 20 (KGs/Units)</b></td><td className={td}>Items where MTD Ach% &gt; 100% — SO exceeding projection. Sorted highest first.</td></tr>
            <tr><td className={td}><b>Bottom 20 (KGs/Units)</b></td><td className={td}>Items where 0 &lt; MTD Ach% &le; 100% — lowest achievement. Sorted lowest first.</td></tr>
            <tr><td className={td}><b>Summary: Item Group / Origin</b></td><td className={td}>Bar charts with aggregated Proj vs Exp vs SO KGs by Item Group or Origin.</td></tr>
            <tr><td className={td}><b>Zero SO</b></td><td className={td}>Projected items with zero sales orders (Proj KGs &gt; 0 but SO KGs = 0).</td></tr>
            <tr><td className={td}><b>Unprojected SO</b></td><td className={td}>SO items with no matching Projection row (unexpected demand).</td></tr>
            <tr><td className={td}><b>Data Audit</b></td><td className={td}>Cleaning rules report + per-item unfiltered vs filtered SO KGs comparison.</td></tr>
          </tbody>
        </table>
      </S>

      <S id="excel" title="10. Excel Export">
        <p>The <b>Download Excel</b> button generates a multi-sheet workbook entirely in the browser:</p>
        <ul className="list-disc pl-5 space-y-1 mt-2">
          <li><b>Config sheet:</b> Days in Month (B4), Days Elapsed (B7) — referenced by all formula cells.</li>
          <li><b>Dashboard sheet:</b> All rows with formula cells for Exp KGs, MTD Ach%, SO vs Proj%, Days Left, Diff, RunRate etc.</li>
          <li><b>Grouped sheets</b> (By New MIS, By Customer Group, By Origin, By Item Name): Pivot-style with parent SUM formulas, row outlines, and expandable groups.</li>
          <li><b>CFA Items sheet:</b> Filtered to 56 CFA items with formulas and TOTAL row.</li>
          <li><b>Summary sheets</b> (Item Group, Origin): Aggregated with TOTAL row and Ach% formulas.</li>
          <li><b>Leaderboard sheets</b> (Top/Bottom 20 KGs/Units): Ranked lists with formulas.</li>
          <li><b>Zero SO / Unprojected SO sheets:</b> Static value sheets.</li>
          <li>All percentage columns formatted as <span className={code}>0.0%</span>. Change Days in Month or Days Elapsed in Config sheet to recalculate all formulas.</li>
        </ul>
      </S>
    </div>
  );
}

type AchSortKey = 'projKg' | 'soKg' | 'achKg' | 'soVsProj' | 'diffKg';

interface AchItemGroup {
  name: string;
  rows: Row[];
  projKg: number; soKg: number; expKg: number;
  achKg: number; soVsProj: number; diffKg: number;
}

function AchAbove100Table({ rows }: { rows: Row[] }) {
  const [expanded, setExpanded] = useState(false);
  const [openItems, setOpenItems] = useState<Set<string>>(new Set());
  const [sortCol, setSortCol] = useState<AchSortKey | null>('achKg');
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('desc');

  // Group the same >100% rows by Item Name so each item can be opened into its
  // customer-wise split. Row selection is untouched — only the presentation.
  const groups = useMemo((): AchItemGroup[] => {
    const map = new Map<string, Row[]>();
    rows.forEach(r => {
      const k = r.itemName || r.itemCode || 'Unknown';
      const arr = map.get(k);
      if (arr) arr.push(r); else map.set(k, [r]);
    });
    return [...map.entries()].map(([name, rs]) => {
      const projKg = rs.reduce((s, r) => s + r.projKg, 0);
      const soKg = rs.reduce((s, r) => s + r.soKg, 0);
      const expKg = rs.reduce((s, r) => s + r.expKg, 0);
      return {
        name, rows: rs, projKg, soKg, expKg,
        achKg: expKg > 0 ? (soKg / expKg) * 100 : 0,
        soVsProj: projKg > 0 ? (soKg / projKg) * 100 : 0,
        diffKg: expKg - soKg,
      };
    });
  }, [rows]);

  const sorted = useMemo(() => {
    if (!sortCol) return groups;
    return [...groups].sort((a, b) => {
      const va = a[sortCol];
      const vb = b[sortCol];
      return sortDir === 'desc' ? vb - va : va - vb;
    });
  }, [groups, sortCol, sortDir]);

  const sortRows = useCallback((rs: Row[]) => {
    if (!sortCol) return rs;
    return [...rs].sort((a, b) => {
      const va = a[sortCol] as number;
      const vb = b[sortCol] as number;
      return sortDir === 'desc' ? vb - va : va - vb;
    });
  }, [sortCol, sortDir]);

  const toggleItem = (name: string) => setOpenItems(prev => {
    const next = new Set(prev);
    if (next.has(name)) next.delete(name); else next.add(name);
    return next;
  });
  const expandAllItems = () => setOpenItems(new Set(groups.map(g => g.name)));
  const collapseAllItems = () => setOpenItems(new Set());

  const toggleSort = (col: AchSortKey) => {
    if (sortCol === col) setSortDir(d => d === 'desc' ? 'asc' : 'desc');
    else { setSortCol(col); setSortDir('desc'); }
  };
  const sIcon = (col: string) => sortCol === col ? (sortDir === 'desc' ? ' ▼' : ' ▲') : ' ⇅';

  return (
    <div>
      <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
        <h3 className="text-sm font-bold text-brand-900">
          MTD Ach% &gt; 100% — {groups.length} items / {rows.length} customer rows
        </h3>
        <div className="flex items-center gap-1.5">
        <button onClick={expandAllItems} className="text-[10px] px-3 py-1.5 border rounded-md hover:bg-gray-100 bg-white shadow-sm">Expand All</button>
        <button onClick={collapseAllItems} className="text-[10px] px-3 py-1.5 border rounded-md hover:bg-gray-100 bg-white shadow-sm">Collapse All</button>
        <button
          onClick={() => setExpanded(e => !e)}
          className={`flex items-center gap-1.5 px-4 py-2 text-xs font-medium rounded-lg border shadow-sm transition-all ${
            expanded
              ? 'bg-brand-900 text-white border-brand-900 hover:bg-brand-800'
              : 'bg-white text-brand-900 border-brand-300 hover:bg-brand-50'
          }`}
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            {expanded
              ? <><polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/><line x1="14" y1="10" x2="21" y2="3"/><line x1="3" y1="21" x2="10" y2="14"/></>
              : <><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></>
            }
          </svg>
          {expanded ? 'Compact View' : 'Full View (Screenshot)'}
        </button>
        </div>
      </div>

      <div className={`bg-white rounded-lg shadow-sm overflow-auto scrollbar-thin ${expanded ? '' : 'max-h-[calc(100vh-340px)]'}`}>
        <table className="w-full text-[11px]">
          <thead className="sticky top-0 z-10">
            <tr>
              <th className="bg-gray-700 text-white px-2 py-2 text-center w-8">#</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-left">Item Name</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('projKg')}>
                Proj KGs{sIcon('projKg')}
              </th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('soKg')}>
                SO KGs{sIcon('soKg')}
              </th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">Exp KGs</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-center cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('achKg')}>
                MTD Ach%{sIcon('achKg')}
              </th>
              <th className="bg-gray-700 text-white px-2 py-2 text-center cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('soVsProj')}>
                SO vs Proj%{sIcon('soVsProj')}
              </th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('diffKg')}>
                Diff KGs{sIcon('diffKg')}
              </th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((g, i) => {
              const isOpen = expanded || openItems.has(g.name);
              return (
                <Fragment key={g.name}>
                  <tr onClick={() => toggleItem(g.name)} className="border-b border-gray-100 hover:bg-blue-50/40 even:bg-gray-50/50 cursor-pointer">
                    <td className="px-2 py-1.5 text-center font-bold text-gray-400">{i + 1}</td>
                    <td className="px-2 py-1.5 max-w-[280px]">
                      <span className="text-brand-700 font-bold mr-1.5">{isOpen ? '▼' : '▶'}</span>
                      <span className="truncate">{g.name}</span>
                      <span className="ml-2 text-[9px] text-gray-400">({g.rows.length} {g.rows.length === 1 ? 'customer' : 'customers'})</span>
                    </td>
                    <td className="px-2 py-1.5 text-right">{fmt(g.projKg)}</td>
                    <td className="px-2 py-1.5 text-right">{fmt(g.soKg)}</td>
                    <td className="px-2 py-1.5 text-right">{fmt(g.expKg)}</td>
                    <td className="px-2 py-1.5 text-center">
                      <span className={`px-2 py-0.5 rounded text-xs font-semibold ${g.achKg > 120 ? 'bg-[#c00000] text-white' : 'bg-[#ff0000] text-white'}`}>
                        {g.achKg.toFixed(1)}%
                      </span>
                    </td>
                    <td className="px-2 py-1.5 text-center">{g.projKg > 0 ? g.soVsProj.toFixed(1) + '%' : '-'}</td>
                    <td className={`px-2 py-1.5 text-right font-semibold ${g.diffKg < 0 ? 'text-red-700' : 'text-green-700'}`}>{fmt(g.diffKg)}</td>
                  </tr>
                  {isOpen && sortRows(g.rows).map((r, j) => (
                    <tr key={`${g.name}-${r.custGroup}-${r.customer}-${r.origin}-${r.itemCode}-${j}`} className="border-b border-gray-50 bg-blue-50/20 hover:bg-yellow-50/50 text-[10px]">
                      <td className="px-2 py-1"></td>
                      <td className="px-2 py-1 pl-8 max-w-[280px] truncate">
                        <span className="font-medium text-gray-800">{r.customer || 'Unknown'}</span>
                        <span className="text-gray-400 ml-2">{r.custGroup}{r.origin ? ` · ${r.origin}` : ''}</span>
                      </td>
                      <td className="px-2 py-1 text-right">{fmt(r.projKg)}</td>
                      <td className="px-2 py-1 text-right">{fmt(r.soKg)}</td>
                      <td className="px-2 py-1 text-right">{fmt(r.expKg)}</td>
                      <td className="px-2 py-1 text-center">
                        <span className={`px-1.5 py-0.5 rounded text-[10px] font-semibold ${r.achKg > 120 ? 'bg-[#c00000] text-white' : 'bg-[#ff0000] text-white'}`}>
                          {r.achKg.toFixed(1)}%
                        </span>
                      </td>
                      <td className="px-2 py-1 text-center">{r.projKg > 0 ? r.soVsProj.toFixed(1) + '%' : '-'}</td>
                      <td className={`px-2 py-1 text-right font-semibold ${r.diffKg < 0 ? 'text-red-700' : 'text-green-700'}`}>{fmt(r.diffKg)}</td>
                    </tr>
                  ))}
                </Fragment>
              );
            })}
            {sorted.length === 0 && (
              <tr><td colSpan={8} className="text-center py-8 text-gray-400">No items with MTD Ach% &gt; 100%</td></tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="mt-2 text-[10px] text-gray-400">
        {expanded ? '📸 Full view — all items expanded to their customer split for screenshot' : `${groups.length} items | ${rows.length} customer rows | Click an item to see its customer-wise breakup`}
      </div>
    </div>
  );
}

function DataAudit({ data, filtered }: { data: DashData; filtered: Row[] }) {
  const [auditSearch, setAuditSearch] = useState('');
  const debouncedAudit = useMemo(() => auditSearch.toLowerCase(), [auditSearch]);

  // Unfiltered per-item audit: sum SO KGs across ALL customer groups (no filter applied)
  const allItemAudit = useMemo(() => {
    const map: Record<string, { itemName: string; itemCode: string; allKg: number; filteredKg: number; interCoKg: number; allRows: number; filteredRows: number }> = {};
    // ALL rows (unfiltered)
    data.rows.forEach(r => {
      const key = r.itemName || r.itemCode;
      if (!map[key]) map[key] = { itemName: r.itemName, itemCode: r.itemCode, allKg: 0, filteredKg: 0, interCoKg: 0, allRows: 0, filteredRows: 0 };
      map[key].allKg += r.soKg;
      map[key].allRows++;
      if (r.custGroup === 'Inter Company') map[key].interCoKg += r.soKg;
    });
    // Filtered rows
    filtered.forEach(r => {
      const key = r.itemName || r.itemCode;
      if (map[key]) {
        map[key].filteredKg += r.soKg;
        map[key].filteredRows++;
      }
    });
    return Object.values(map)
      .filter(r => r.allKg > 0)
      .map(r => ({ ...r, allKg: Math.round(r.allKg * 100) / 100, filteredKg: Math.round(r.filteredKg * 100) / 100, interCoKg: Math.round(r.interCoKg * 100) / 100, diff: Math.round((r.allKg - r.filteredKg) * 100) / 100 }))
      .sort((a, b) => b.diff - a.diff);
  }, [data.rows, filtered]);

  const filteredAudit = debouncedAudit
    ? allItemAudit.filter(r => r.itemName.toLowerCase().includes(debouncedAudit) || r.itemCode.toLowerCase().includes(debouncedAudit))
    : allItemAudit;

  return (
    <div className="space-y-6">
      {/* Cleaning Rules Report */}
      <div className="bg-white rounded-lg shadow-sm p-5">
        <h3 className="text-sm font-bold text-brand-900 mb-3">SO Cleaning Rules Report</h3>
        <p className="text-xs text-gray-500 mb-3">Rows removed from raw SO data before processing. Total raw: <b>{data.config.soRowsRaw}</b> → After cleaning: <b>{data.config.soRowsClean}</b> ({data.config.soRowsRaw - data.config.soRowsClean} removed)</p>
        <table className="w-full text-xs border-collapse">
          <thead>
            <tr className="bg-gray-100">
              <th className="text-left px-3 py-2 border">Cleaning Rule</th>
              <th className="text-right px-3 py-2 border">Rows Removed</th>
              <th className="text-right px-3 py-2 border">Stock Qty (KGs) Removed</th>
            </tr>
          </thead>
          <tbody>
            {data.cleaningStats.length === 0 && (
              <tr><td colSpan={3} className="px-3 py-2 border text-center text-gray-400">No rows removed by cleaning</td></tr>
            )}
            {data.cleaningStats.map(s => (
              <tr key={s.rule} className="hover:bg-gray-50">
                <td className="px-3 py-2 border font-medium">{s.rule}</td>
                <td className="px-3 py-2 border text-right">{s.removedRows.toLocaleString('en-IN')}</td>
                <td className="px-3 py-2 border text-right text-red-600 font-medium">{fmt(s.removedKg)}</td>
              </tr>
            ))}
            <tr className="bg-gray-100 font-bold">
              <td className="px-3 py-2 border">TOTAL</td>
              <td className="px-3 py-2 border text-right">{(data.config.soRowsRaw - data.config.soRowsClean).toLocaleString('en-IN')}</td>
              <td className="px-3 py-2 border text-right text-red-600">{fmt(data.cleaningStats.reduce((a, s) => a + s.removedKg, 0))}</td>
            </tr>
          </tbody>
        </table>
      </div>

      {/* Per-Item SO KGs Audit */}
      <div className="bg-white rounded-lg shadow-sm p-5">
        <h3 className="text-sm font-bold text-brand-900 mb-1">Per-Item SO KGs Audit — Unfiltered vs Filtered</h3>
        <p className="text-xs text-gray-500 mb-3">Compare SO KGs across all customer groups (unfiltered) vs your current filter selection. Diff = KGs excluded by filters (e.g., Inter Company unchecked).</p>
        <input
          value={auditSearch}
          onChange={e => setAuditSearch(e.target.value)}
          placeholder="Search item name or code..."
          className="px-3 py-1.5 border rounded-md text-xs w-80 outline-none focus:border-brand-500 focus:ring-1 focus:ring-brand-500/20 mb-3"
        />
        <div className="max-h-[500px] overflow-auto">
          <table className="w-full text-xs border-collapse">
            <thead className="sticky top-0">
              <tr className="bg-gray-100">
                <th className="text-left px-3 py-2 border">Item Name</th>
                <th className="text-right px-3 py-2 border">All SO KGs</th>
                <th className="text-right px-3 py-2 border">Inter Co. KGs</th>
                <th className="text-right px-3 py-2 border">Filtered SO KGs</th>
                <th className="text-right px-3 py-2 border">Diff (Excluded)</th>
                <th className="text-right px-3 py-2 border">All Rows</th>
                <th className="text-right px-3 py-2 border">Filtered Rows</th>
              </tr>
            </thead>
            <tbody>
              {filteredAudit.slice(0, 200).map(r => (
                <tr key={r.itemName} className={`hover:bg-gray-50 ${r.diff > 0 ? 'bg-yellow-50' : ''}`}>
                  <td className="px-3 py-1.5 border font-medium max-w-xs truncate" title={r.itemName}>{r.itemName}</td>
                  <td className="px-3 py-1.5 border text-right font-mono">{fmt(r.allKg)}</td>
                  <td className="px-3 py-1.5 border text-right font-mono text-blue-600">{fmt(r.interCoKg)}</td>
                  <td className="px-3 py-1.5 border text-right font-mono">{fmt(r.filteredKg)}</td>
                  <td className={`px-3 py-1.5 border text-right font-mono font-bold ${r.diff > 0 ? 'text-red-600' : 'text-green-600'}`}>{r.diff > 0 ? '+' : ''}{fmt(r.diff)}</td>
                  <td className="px-3 py-1.5 border text-right">{r.allRows}</td>
                  <td className="px-3 py-1.5 border text-right">{r.filteredRows}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

export default function Dashboard({ data, onReUpload }: DashboardProps) {
  const [tab, setTab] = useState<Tab>('dashboard');
  const [search, setSearch] = useState('');
  const [cgFilter, setCgFilter] = useState<Set<string>>(() => new Set(data.filters.custGroups.filter(c => c !== 'Inter Company')));
  const [originFilter, setOriginFilter] = useState<Set<string>>(() => new Set(data.filters.origins));
  const [igFilter, setIgFilter] = useState<Set<string>>(() => new Set(data.filters.itemGroups));
  const searchTimer = useRef<ReturnType<typeof setTimeout>>();
  const [debouncedSearch, setDebouncedSearch] = useState('');

  const onSearch = useCallback((v: string) => {
    setSearch(v);
    if (searchTimer.current) clearTimeout(searchTimer.current);
    searchTimer.current = setTimeout(() => setDebouncedSearch(v.toLowerCase()), 300);
  }, []);

  const filtered = useMemo(() => {
    if (!data) return [];
    return data.rows.filter(r => {
      if (!cgFilter.has(r.custGroup)) return false;
      if (!originFilter.has(r.origin)) return false;
      if (r.itemGroup && !igFilter.has(r.itemGroup)) return false;
      if (debouncedSearch) {
        const s = debouncedSearch;
        if (!(
          r.itemName?.toLowerCase().includes(s) ||
          r.itemCode?.toLowerCase().includes(s) ||
          r.customer?.toLowerCase().includes(s) ||
          r.custGroup?.toLowerCase().includes(s) ||
          r.newMIS?.toLowerCase().includes(s)
        )) return false;
      }
      return true;
    });
  }, [data, cgFilter, originFilter, igFilter, debouncedSearch]);

  // Dynamic summaries & leaderboards — respect current filter state
  const dynSummaryIG = useMemo(() => aggGroup(filtered, 'itemGroup'), [filtered]);
  const dynSummaryOrigin = useMemo(() => aggGroup(filtered, 'origin'), [filtered]);

  const dynLeaderboards = useMemo(() => {
    const withProj = filtered.filter(r => r.isProj && r.expKg > 0);
    return {
      top20Kg: [...withProj].sort((a, b) => b.achKg - a.achKg).filter(r => r.achKg > 100).slice(0, 20),
      top20Units: [...withProj].sort((a, b) => b.achUnits - a.achUnits).filter(r => r.achUnits > 100).slice(0, 20),
      bot20Kg: [...withProj].sort((a, b) => a.achKg - b.achKg).filter(r => r.achKg > 0 && r.achKg <= 100).slice(0, 20),
      bot20Units: [...withProj].sort((a, b) => a.achUnits - b.achUnits).filter(r => r.achUnits > 0 && r.achUnits <= 100).slice(0, 20),
    };
  }, [filtered]);

  const cfg = data.config;

  return (
    <div className="min-h-screen">
      {/* Header */}
      <header className="bg-gradient-to-r from-brand-900 to-brand-700 text-white px-6 py-3 flex items-center justify-between sticky top-0 z-50 shadow-lg">
        <h1 className="text-lg font-bold">Farmley — Run Rate Dashboard</h1>
        <div className="flex items-center gap-4 text-[11px]">
          <span className="bg-white/15 px-3 py-1 rounded-full">Days: <b>{cfg.daysInMonth}</b></span>
          <span className="bg-white/15 px-3 py-1 rounded-full">SO: <b>{cfg.dateFrom}</b> → <b>{cfg.dateTo}</b></span>
          <span className="bg-white/15 px-3 py-1 rounded-full">Elapsed: <b>{cfg.daysElapsed}d</b></span>
          <span className="bg-white/15 px-3 py-1 rounded-full">Rows: <b>{cfg.totalRows}</b></span>
          <button
            onClick={() => exportDashboardExcel(data, CFA_ITEM_NAMES)}
            className="flex items-center gap-1.5 bg-white/20 hover:bg-white/30 px-4 py-1.5 rounded-full font-semibold transition-all text-[11px]"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download Excel
          </button>
          <button
            onClick={onReUpload}
            className="flex items-center gap-1.5 bg-white/20 hover:bg-white/30 px-4 py-1.5 rounded-full font-semibold transition-all text-[11px]"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            Re-upload Files
          </button>
        </div>
      </header>

      {/* Tabs */}
      <div className="bg-white border-b shadow-sm sticky top-[52px] z-40 overflow-x-auto scrollbar-thin">
        <div className="flex px-4">
          {TABS.map(t => (
            <button
              key={t.id}
              onClick={() => setTab(t.id)}
              className={`px-4 py-2.5 text-xs font-medium whitespace-nowrap border-b-2 transition-all ${
                tab === t.id ? 'text-brand-900 border-brand-900 font-semibold' : 'text-gray-500 border-transparent hover:text-brand-700 hover:bg-gray-50'
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>
      </div>

      <div className="p-4 max-w-[1920px] mx-auto">
        {/* Filters */}
        {['ach-above-100', 'dashboard', 'cfa-items', 'new-mis', 'cust-group', 'origin', 'item-name', 'top-kg', 'top-unit', 'bot-kg', 'bot-unit', 'summary-ig', 'summary-origin', 'zero-so', 'unproj'].includes(tab) && (
          <div className="bg-white rounded-lg shadow-sm p-4 mb-4 flex gap-4 flex-wrap">
            <FilterBar label="Customer Group" options={data.filters.custGroups} selected={cgFilter} onChange={setCgFilter} />
            <FilterBar label="Origin" options={data.filters.origins} selected={originFilter} onChange={setOriginFilter} />
            <FilterBar label="Item Group" options={data.filters.itemGroups} selected={igFilter} onChange={setIgFilter} />
          </div>
        )}

        {/* Legend + Search */}
        {['ach-above-100', 'dashboard', 'cfa-items', 'new-mis', 'cust-group', 'origin', 'item-name'].includes(tab) && (
          <>
            <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
              <div className="flex gap-2 flex-wrap text-[10px]">
                <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-[#c00000]"></span> &gt;120% Urgent</span>
                <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-[#ff0000]"></span> 100-120%</span>
                <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-[#ffc000]"></span> 80-100%</span>
                <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-[#92d050]"></span> 50-80%</span>
                <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-[#00b050]"></span> &lt;50%</span>
                <span className="flex items-center gap-1"><span className="w-3 h-3 rounded bg-[#d9d9d9]"></span> No SO</span>
              </div>
              <input
                value={search}
                onChange={e => onSearch(e.target.value)}
                placeholder="Search item, code, customer..."
                className="px-3 py-1.5 border rounded-md text-xs w-64 outline-none focus:border-brand-500 focus:ring-1 focus:ring-brand-500/20"
              />
            </div>
            <StatsCards rows={tab === 'cfa-items' ? filtered.filter(r => CFA_ITEM_NAMES.has(r.itemName)) : tab === 'ach-above-100' ? filtered.filter(r => r.achKg > 100) : filtered} />
          </>
        )}

        {/* Tab Content */}
        {tab === 'ach-above-100' && <AchAbove100Table rows={filtered.filter(r => r.achKg > 100)} />}
        {tab === 'instructions' && <InstructionsPanel cfg={cfg} />}
        {tab === 'dashboard' && <DataTable data={filtered} />}
        {tab === 'cfa-items' && <GroupedView rows={filtered.filter(r => CFA_ITEM_NAMES.has(r.itemName))} groupBy="itemName" subGroupBy="customer" daysElapsed={cfg.daysElapsed} />}
        {tab === 'new-mis' && <GroupedView rows={filtered} groupBy="newMIS" daysElapsed={cfg.daysElapsed} />}
        {tab === 'cust-group' && <GroupedView rows={filtered} groupBy="custGroup" subGroupBy="customer" daysElapsed={cfg.daysElapsed} />}
        {tab === 'origin' && <GroupedView rows={filtered} groupBy="origin" daysElapsed={cfg.daysElapsed} />}
        {tab === 'item-name' && <GroupedView rows={filtered} groupBy="itemName" subGroupBy="customer" daysElapsed={cfg.daysElapsed} />}
        {tab === 'top-kg' && <Leaderboard data={dynLeaderboards.top20Kg} title="Top 20 Exceeding — KGs" subtitle="SKUs where SO exceeds projection pace — procurement action needed" />}
        {tab === 'top-unit' && <Leaderboard data={dynLeaderboards.top20Units} title="Top 20 Exceeding — Units" subtitle="SKUs where SO exceeds projection pace by unit count" />}
        {tab === 'bot-kg' && <Leaderboard data={dynLeaderboards.bot20Kg} title="Bottom 20 Lagging — KGs" subtitle="SKUs with lowest SO achievement — demand visibility needed" />}
        {tab === 'bot-unit' && <Leaderboard data={dynLeaderboards.bot20Units} title="Bottom 20 Lagging — Units" subtitle="SKUs with lowest SO achievement by unit count" />}
        {tab === 'summary-ig' && <SummaryChart data={dynSummaryIG} title="Summary by Item Group" />}
        {tab === 'summary-origin' && <SummaryChart data={dynSummaryOrigin} title="Summary by Origin" />}
        {tab === 'zero-so' && (
          <div>
            <h3 className="text-sm font-bold text-brand-900 mb-2">Projected Items with Zero Sales Orders ({filtered.filter(r => r.isProj && r.soKg === 0 && r.soUnits === 0 && r.projKg > 0).length} items)</h3>
            <DataTable data={filtered.filter(r => r.isProj && r.soKg === 0 && r.soUnits === 0 && r.projKg > 0)} />
          </div>
        )}
        {tab === 'unproj' && (
          <div>
            <h3 className="text-sm font-bold text-brand-900 mb-2">Unprojected SO Items ({filtered.filter(r => r.isUnproj).length} items)</h3>
            <DataTable data={filtered.filter(r => r.isUnproj)} />
          </div>
        )}
        {tab === 'data-audit' && <DataAudit data={data} filtered={filtered} />}
      </div>
    </div>
  );
}
