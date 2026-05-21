'use client';
import { useState, useMemo, useCallback, useRef } from 'react';
import { Row, achClass, fmt, fmtPct } from '@/lib/helpers';

type GroupKey = 'itemGroup' | 'newMIS' | 'custGroup' | 'origin' | 'itemName' | 'customer';

interface Props {
  rows: Row[];
  groupBy: GroupKey;
  subGroupBy?: GroupKey;
}

interface GroupAgg {
  name: string;
  projKg: number; projUnits: number; expKg: number; expUnits: number;
  soKg: number; soUnits: number; achKg: number; achUnits: number;
  diffKg: number; diffUnits: number; count: number;
  children: Row[];
  subGroups?: { name: string; rows: Row[]; agg: Omit<GroupAgg, 'name' | 'children' | 'subGroups'> }[];
}

function aggregate(rows: Row[]) {
  const projKg = rows.reduce((s, r) => s + r.projKg, 0);
  const projUnits = rows.reduce((s, r) => s + r.projUnits, 0);
  const expKg = rows.reduce((s, r) => s + r.expKg, 0);
  const expUnits = rows.reduce((s, r) => s + r.expUnits, 0);
  const soKg = rows.reduce((s, r) => s + r.soKg, 0);
  const soUnits = rows.reduce((s, r) => s + r.soUnits, 0);
  return {
    projKg, projUnits, expKg, expUnits, soKg, soUnits,
    achKg: expKg > 0 ? (soKg / expKg) * 100 : 0,
    achUnits: expUnits > 0 ? (soUnits / expUnits) * 100 : 0,
    diffKg: expKg - soKg,
    diffUnits: expUnits - soUnits,
    count: rows.length,
  };
}

export default function GroupedView({ rows, groupBy, subGroupBy }: Props) {
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [expandedSub, setExpandedSub] = useState<Set<string>>(new Set());
  const [groupSearch, setGroupSearch] = useState('');
  const [debouncedGS, setDebouncedGS] = useState('');
  const gsTimer = useRef<ReturnType<typeof setTimeout>>();

  const onGroupSearch = useCallback((v: string) => {
    setGroupSearch(v);
    if (gsTimer.current) clearTimeout(gsTimer.current);
    gsTimer.current = setTimeout(() => {
      const s = v.toLowerCase();
      setDebouncedGS(s);
      // Auto-expand matching groups when searching
      if (s) {
        setExpanded(prev => {
          const next = new Set(prev);
          // Will be handled in filteredGroups
          return next;
        });
      }
    }, 300);
  }, []);

  const groups = useMemo((): GroupAgg[] => {
    const map: Record<string, Row[]> = {};
    rows.forEach(r => {
      const k = (r[groupBy] as string) || 'Unknown';
      if (!map[k]) map[k] = [];
      map[k].push(r);
    });
    return Object.entries(map).map(([name, children]) => {
      const agg = aggregate(children);
      let subGroups: GroupAgg['subGroups'];
      if (subGroupBy) {
        const subMap: Record<string, Row[]> = {};
        children.forEach(r => {
          const sk = (r[subGroupBy] as string) || 'Unknown';
          if (!subMap[sk]) subMap[sk] = [];
          subMap[sk].push(r);
        });
        subGroups = Object.entries(subMap).map(([sn, sr]) => ({ name: sn, rows: sr, agg: aggregate(sr) }))
          .sort((a, b) => b.agg.achKg - a.agg.achKg);
      }
      return { name, ...agg, children, subGroups };
    }).sort((a, b) => b.achKg - a.achKg);
  }, [rows, groupBy, subGroupBy]);

  const filteredGroups = useMemo(() => {
    if (!debouncedGS) return groups;
    return groups.filter(g => {
      // Match on group name
      if (g.name.toLowerCase().includes(debouncedGS)) return true;
      // Match on any child item code/name/customer
      return g.children.some(r =>
        r.itemCode?.toLowerCase().includes(debouncedGS) ||
        r.itemName?.toLowerCase().includes(debouncedGS) ||
        r.customer?.toLowerCase().includes(debouncedGS)
      );
    });
  }, [groups, debouncedGS]);

  // Auto-expand all groups when searching
  const effectiveExpanded = useMemo(() => {
    if (debouncedGS) {
      return new Set(filteredGroups.map(g => g.name));
    }
    return expanded;
  }, [debouncedGS, filteredGroups, expanded]);

  const toggleGroup = useCallback((name: string) => {
    if (debouncedGS) return; // Don't toggle during search
    setExpanded(prev => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name); else next.add(name);
      return next;
    });
  }, [debouncedGS]);

  const toggleSub = useCallback((key: string) => {
    setExpandedSub(prev => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key); else next.add(key);
      return next;
    });
  }, []);

  const expandAll = useCallback(() => {
    setExpanded(new Set(groups.map(g => g.name)));
  }, [groups]);

  const collapseAll = useCallback(() => {
    setExpanded(new Set());
    setExpandedSub(new Set());
  }, []);

  return (
    <div>
      {/* Search + controls bar */}
      <div className="flex items-center justify-between mb-3 gap-3 flex-wrap">
        <div className="flex items-center gap-2">
          <div className="relative">
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400">
              <circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>
            </svg>
            <input
              value={groupSearch}
              onChange={e => onGroupSearch(e.target.value)}
              placeholder="Search item name, code, customer..."
              className="pl-9 pr-3 py-2 border rounded-lg text-xs w-80 outline-none focus:border-brand-500 focus:ring-2 focus:ring-brand-500/20 bg-white shadow-sm"
            />
            {groupSearch && (
              <button
                onClick={() => { setGroupSearch(''); setDebouncedGS(''); }}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-sm"
              >✕</button>
            )}
          </div>
          <span className="text-[11px] text-gray-500">
            {filteredGroups.length} / {groups.length} groups
          </span>
        </div>
        <div className="flex gap-1">
          <button onClick={expandAll} className="text-[10px] px-3 py-1.5 border rounded-md hover:bg-gray-100 bg-white shadow-sm">Expand All</button>
          <button onClick={collapseAll} className="text-[10px] px-3 py-1.5 border rounded-md hover:bg-gray-100 bg-white shadow-sm">Collapse All</button>
        </div>
      </div>

      {/* Table */}
      <div className="bg-white rounded-lg shadow-sm overflow-auto max-h-[calc(100vh-340px)] scrollbar-thin">
        <table className="w-full text-[11px]">
          <thead className="sticky top-0 z-10">
            <tr>
              <th className="bg-gray-700 text-white px-3 py-2 text-left w-8"></th>
              <th className="bg-gray-700 text-white px-3 py-2 text-left min-w-[200px]">Group</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">Count</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">Proj KGs</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">Proj Units</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">Exp KGs</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">SO KGs</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">SO Units</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-center">Ach% KGs</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-center">Ach% Units</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">Diff KGs</th>
              <th className="bg-gray-700 text-white px-2 py-2 text-right">Diff Units</th>
            </tr>
          </thead>
          <tbody>
            {filteredGroups.map(g => {
              const isOpen = effectiveExpanded.has(g.name);
              return (
                <GroupRows key={g.name} g={g} isOpen={isOpen} toggleGroup={toggleGroup}
                  subGroupBy={subGroupBy} expandedSub={expandedSub} toggleSub={toggleSub}
                  highlight={debouncedGS} />
              );
            })}
            {filteredGroups.length === 0 && (
              <tr><td colSpan={12} className="text-center py-8 text-gray-400">No groups match your search</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function HighlightText({ text, highlight }: { text: string; highlight: string }) {
  if (!highlight || !text) return <>{text}</>;
  const idx = text.toLowerCase().indexOf(highlight);
  if (idx === -1) return <>{text}</>;
  return (
    <>
      {text.slice(0, idx)}
      <mark className="bg-yellow-200 rounded px-0.5">{text.slice(idx, idx + highlight.length)}</mark>
      {text.slice(idx + highlight.length)}
    </>
  );
}

function GroupRows({ g, isOpen, toggleGroup, subGroupBy, expandedSub, toggleSub, highlight }: {
  g: GroupAgg; isOpen: boolean; toggleGroup: (n: string) => void;
  subGroupBy?: GroupKey; expandedSub: Set<string>; toggleSub: (k: string) => void;
  highlight: string;
}) {
  return (
    <>
      <tr
        onClick={() => toggleGroup(g.name)}
        className="cursor-pointer hover:bg-blue-50 border-b border-gray-200 bg-gray-50"
      >
        <td className="px-3 py-2 text-center font-bold text-brand-700">{isOpen ? '▼' : '▶'}</td>
        <td className="px-3 py-2 font-semibold text-brand-900">
          <HighlightText text={g.name} highlight={highlight} />
        </td>
        <td className="px-2 py-2 text-right text-gray-500">{g.count}</td>
        <td className="px-2 py-2 text-right">{fmt(g.projKg)}</td>
        <td className="px-2 py-2 text-right">{fmt(g.projUnits, 0)}</td>
        <td className="px-2 py-2 text-right">{fmt(g.expKg)}</td>
        <td className="px-2 py-2 text-right">{fmt(g.soKg)}</td>
        <td className="px-2 py-2 text-right">{fmt(g.soUnits, 0)}</td>
        <td className="px-2 py-2 text-center"><span className={`px-2 py-0.5 rounded text-xs ${achClass(g.achKg)}`}>{fmtPct(g.achKg)}</span></td>
        <td className="px-2 py-2 text-center"><span className={`px-2 py-0.5 rounded text-xs ${achClass(g.achUnits)}`}>{fmtPct(g.achUnits)}</span></td>
        <td className={`px-2 py-2 text-right ${g.diffKg < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(g.diffKg)}</td>
        <td className={`px-2 py-2 text-right ${g.diffUnits < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(g.diffUnits, 0)}</td>
      </tr>
      {isOpen && (
        subGroupBy && g.subGroups ? (
          g.subGroups.map(sg => {
            const subKey = `${g.name}|${sg.name}`;
            const subOpen = expandedSub.has(subKey);
            return (
              <SubGroupRows key={subKey} sg={sg} subKey={subKey} subOpen={subOpen} toggleSub={toggleSub} highlight={highlight} />
            );
          })
        ) : (
          g.children.map(r => <ItemRow key={`${r.itemCode}-${r.customer}-${r.origin}`} r={r} indent={1} highlight={highlight} />)
        )
      )}
    </>
  );
}

function SubGroupRows({ sg, subKey, subOpen, toggleSub, highlight }: {
  sg: { name: string; rows: Row[]; agg: any }; subKey: string; subOpen: boolean; toggleSub: (k: string) => void;
  highlight: string;
}) {
  return (
    <>
      <tr onClick={() => toggleSub(subKey)} className="cursor-pointer hover:bg-blue-50/60 border-b border-gray-100 bg-blue-50/30">
        <td className="px-3 py-1.5 text-center text-gray-400 pl-6">{subOpen ? '▾' : '▸'}</td>
        <td className="px-3 py-1.5 pl-8 font-medium text-gray-700">
          <HighlightText text={sg.name} highlight={highlight} />
        </td>
        <td className="px-2 py-1.5 text-right text-gray-400">{sg.rows.length}</td>
        <td className="px-2 py-1.5 text-right">{fmt(sg.agg.projKg)}</td>
        <td className="px-2 py-1.5 text-right">{fmt(sg.agg.projUnits, 0)}</td>
        <td className="px-2 py-1.5 text-right">{fmt(sg.agg.expKg)}</td>
        <td className="px-2 py-1.5 text-right">{fmt(sg.agg.soKg)}</td>
        <td className="px-2 py-1.5 text-right">{fmt(sg.agg.soUnits, 0)}</td>
        <td className="px-2 py-1.5 text-center"><span className={`px-2 py-0.5 rounded text-xs ${achClass(sg.agg.achKg)}`}>{fmtPct(sg.agg.achKg)}</span></td>
        <td className="px-2 py-1.5 text-center"><span className={`px-2 py-0.5 rounded text-xs ${achClass(sg.agg.achUnits)}`}>{fmtPct(sg.agg.achUnits)}</span></td>
        <td className={`px-2 py-1.5 text-right ${sg.agg.diffKg < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(sg.agg.diffKg)}</td>
        <td className={`px-2 py-1.5 text-right ${sg.agg.diffUnits < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(sg.agg.diffUnits, 0)}</td>
      </tr>
      {subOpen && sg.rows.map(r => <ItemRow key={`${r.itemCode}-${r.customer}-${r.origin}`} r={r} indent={2} highlight={highlight} />)}
    </>
  );
}

function ItemRow({ r, indent, highlight }: { r: Row; indent: number; highlight: string }) {
  return (
    <tr className="border-b border-gray-50 hover:bg-yellow-50/40 text-[10px]">
      <td className="px-3 py-1"></td>
      <td className="py-1 text-gray-600" style={{ paddingLeft: 12 + indent * 20 }}>
        <span className="font-medium text-gray-800">
          <HighlightText text={r.itemCode} highlight={highlight} />
        </span>
        <span className="text-gray-400 ml-2">
          <HighlightText text={r.itemName?.slice(0, 50) || ''} highlight={highlight} />
        </span>
      </td>
      <td className="px-2 py-1 text-right text-gray-400">
        <HighlightText text={r.customer?.slice(0, 25) || ''} highlight={highlight} />
      </td>
      <td className="px-2 py-1 text-right">{fmt(r.projKg)}</td>
      <td className="px-2 py-1 text-right">{fmt(r.projUnits, 0)}</td>
      <td className="px-2 py-1 text-right">{fmt(r.expKg)}</td>
      <td className="px-2 py-1 text-right">{fmt(r.soKg)}</td>
      <td className="px-2 py-1 text-right">{fmt(r.soUnits, 0)}</td>
      <td className="px-2 py-1 text-center"><span className={`px-1.5 py-0.5 rounded text-[10px] ${achClass(r.achKg)}`}>{fmtPct(r.achKg)}</span></td>
      <td className="px-2 py-1 text-center"><span className={`px-1.5 py-0.5 rounded text-[10px] ${achClass(r.achUnits)}`}>{fmtPct(r.achUnits)}</span></td>
      <td className={`px-2 py-1 text-right ${r.diffKg < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(r.diffKg)}</td>
      <td className={`px-2 py-1 text-right ${r.diffUnits < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(r.diffUnits, 0)}</td>
    </tr>
  );
}
