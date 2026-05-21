'use client';
import { useState, useMemo, useCallback, useRef } from 'react';
import { DashData, Row, fmt } from '@/lib/helpers';
import { exportDashboardExcel } from '@/lib/exportExcel';
import { aggGroup } from '@/lib/processData';
import StatsCards from './StatsCards';
import FilterBar from './FilterBar';
import DataTable from './DataTable';
import GroupedView from './GroupedView';
import Leaderboard from './Leaderboard';
import SummaryChart from './SummaryChart';

type Tab = 'dashboard' | 'cfa-items' | 'new-mis' | 'cust-group' | 'origin' | 'item-name' | 'top-kg' | 'top-unit' | 'bot-kg' | 'bot-unit' | 'summary-ig' | 'summary-origin' | 'zero-so' | 'unproj' | 'data-audit';

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
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'cfa-items', label: 'CFA Items Run Rate' },
  { id: 'new-mis', label: 'By New MIS' },
  { id: 'cust-group', label: 'By Customer Group' },
  { id: 'origin', label: 'By Origin' },
  { id: 'item-name', label: 'By Item Name' },
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
        {['dashboard', 'cfa-items', 'new-mis', 'cust-group', 'origin', 'item-name', 'top-kg', 'top-unit', 'bot-kg', 'bot-unit', 'summary-ig', 'summary-origin', 'zero-so', 'unproj'].includes(tab) && (
          <div className="bg-white rounded-lg shadow-sm p-4 mb-4 flex gap-4 flex-wrap">
            <FilterBar label="Customer Group" options={data.filters.custGroups} selected={cgFilter} onChange={setCgFilter} />
            <FilterBar label="Origin" options={data.filters.origins} selected={originFilter} onChange={setOriginFilter} />
            <FilterBar label="Item Group" options={data.filters.itemGroups} selected={igFilter} onChange={setIgFilter} />
          </div>
        )}

        {/* Legend + Search */}
        {['dashboard', 'cfa-items', 'new-mis', 'cust-group', 'origin', 'item-name'].includes(tab) && (
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
            <StatsCards rows={tab === 'cfa-items' ? filtered.filter(r => CFA_ITEM_NAMES.has(r.itemName)) : filtered} />
          </>
        )}

        {/* Tab Content */}
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
