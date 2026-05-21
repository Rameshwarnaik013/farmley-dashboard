'use client';
import { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { DashData, Row } from '@/lib/helpers';
import StatsCards from './StatsCards';
import FilterBar from './FilterBar';
import DataTable from './DataTable';
import GroupedView from './GroupedView';
import Leaderboard from './Leaderboard';
import SummaryChart from './SummaryChart';

type Tab = 'dashboard' | 'item-group' | 'new-mis' | 'cust-group' | 'origin' | 'item-name' | 'top-kg' | 'top-unit' | 'bot-kg' | 'bot-unit' | 'summary-ig' | 'summary-origin' | 'zero-so' | 'unproj';

const TABS: { id: Tab; label: string }[] = [
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'item-group', label: 'By Item Group' },
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
];

export default function Dashboard() {
  const [data, setData] = useState<DashData | null>(null);
  const [tab, setTab] = useState<Tab>('dashboard');
  const [search, setSearch] = useState('');
  const [cgFilter, setCgFilter] = useState<Set<string>>(new Set());
  const [originFilter, setOriginFilter] = useState<Set<string>>(new Set());
  const [igFilter, setIgFilter] = useState<Set<string>>(new Set());
  const searchTimer = useRef<ReturnType<typeof setTimeout>>();
  const [debouncedSearch, setDebouncedSearch] = useState('');

  useEffect(() => {
    fetch('/data.json')
      .then(r => r.json())
      .then((d: DashData) => {
        setData(d);
        setCgFilter(new Set(d.filters.custGroups.filter(c => c !== 'Inter Company')));
        setOriginFilter(new Set(d.filters.origins));
        setIgFilter(new Set(d.filters.itemGroups));
      });
  }, []);

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

  if (!data) {
    return (
      <div className="flex items-center justify-center h-screen">
        <div className="text-center">
          <div className="w-12 h-12 border-4 border-brand-500 border-t-transparent rounded-full animate-spin mx-auto mb-4"></div>
          <p className="text-gray-500">Loading dashboard data...</p>
        </div>
      </div>
    );
  }

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
          <a
            href="/api/download"
            className="flex items-center gap-1.5 bg-white/20 hover:bg-white/30 px-4 py-1.5 rounded-full font-semibold transition-all text-[11px]"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download Excel
          </a>
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
        {['dashboard', 'item-group', 'new-mis', 'cust-group', 'origin', 'item-name', 'zero-so', 'unproj'].includes(tab) && (
          <div className="bg-white rounded-lg shadow-sm p-4 mb-4 flex gap-4 flex-wrap">
            <FilterBar label="Customer Group" options={data.filters.custGroups} selected={cgFilter} onChange={setCgFilter} />
            <FilterBar label="Origin" options={data.filters.origins} selected={originFilter} onChange={setOriginFilter} />
            <FilterBar label="Item Group" options={data.filters.itemGroups} selected={igFilter} onChange={setIgFilter} />
          </div>
        )}

        {/* Legend + Search */}
        {['dashboard', 'item-group', 'new-mis', 'cust-group', 'origin', 'item-name'].includes(tab) && (
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
            <StatsCards rows={filtered} />
          </>
        )}

        {/* Tab Content */}
        {tab === 'dashboard' && <DataTable data={filtered} />}
        {tab === 'item-group' && <GroupedView rows={filtered} groupBy="itemGroup" />}
        {tab === 'new-mis' && <GroupedView rows={filtered} groupBy="newMIS" />}
        {tab === 'cust-group' && <GroupedView rows={filtered} groupBy="custGroup" subGroupBy="customer" />}
        {tab === 'origin' && <GroupedView rows={filtered} groupBy="origin" />}
        {tab === 'item-name' && <GroupedView rows={filtered} groupBy="itemName" subGroupBy="customer" />}
        {tab === 'top-kg' && <Leaderboard data={data.top20Kg} title="Top 20 Exceeding — KGs" subtitle="SKUs where SO exceeds projection pace — procurement action needed" />}
        {tab === 'top-unit' && <Leaderboard data={data.top20Units} title="Top 20 Exceeding — Units" subtitle="SKUs where SO exceeds projection pace by unit count" />}
        {tab === 'bot-kg' && <Leaderboard data={data.bot20Kg} title="Bottom 20 Lagging — KGs" subtitle="SKUs with lowest SO achievement — demand visibility needed" />}
        {tab === 'bot-unit' && <Leaderboard data={data.bot20Units} title="Bottom 20 Lagging — Units" subtitle="SKUs with lowest SO achievement by unit count" />}
        {tab === 'summary-ig' && <SummaryChart data={data.summaryIG} title="Summary by Item Group" />}
        {tab === 'summary-origin' && <SummaryChart data={data.summaryOrigin} title="Summary by Origin" />}
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
      </div>
    </div>
  );
}
