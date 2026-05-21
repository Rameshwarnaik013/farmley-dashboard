'use client';
import { useState, useMemo } from 'react';
import { Row, achClass, fmt, fmtPct } from '@/lib/helpers';

type SortKey = 'projKg' | 'soKg' | 'achKg' | null;

export default function Leaderboard({ data, title, subtitle }: { data: Row[]; title: string; subtitle: string }) {
  const [sortCol, setSortCol] = useState<SortKey>(null);
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('desc');
  const [fullView, setFullView] = useState(false);

  const sorted = useMemo(() => {
    if (!sortCol) return data;
    return [...data].sort((a, b) => {
      const va = a[sortCol] as number;
      const vb = b[sortCol] as number;
      return sortDir === 'desc' ? vb - va : va - vb;
    });
  }, [data, sortCol, sortDir]);

  const toggleSort = (col: NonNullable<SortKey>) => {
    if (sortCol === col) setSortDir(d => d === 'desc' ? 'asc' : 'desc');
    else { setSortCol(col); setSortDir('desc'); }
  };
  const sIcon = (col: string) => sortCol === col ? (sortDir === 'desc' ? ' ▼' : ' ▲') : ' ⇅';

  return (
    <div className="bg-white rounded-lg shadow-sm overflow-hidden">
      <div className="px-4 py-3 border-b bg-gray-50 flex items-center justify-between">
        <div>
          <h3 className="text-sm font-bold text-brand-900">{title}</h3>
          <p className="text-[10px] text-gray-500">{subtitle}</p>
        </div>
        <button
          onClick={() => setFullView(v => !v)}
          className={`flex items-center gap-1.5 text-[10px] px-3 py-1.5 border rounded-md shadow-sm transition-all ${
            fullView ? 'bg-brand-900 text-white border-brand-900 hover:bg-brand-800' : 'bg-white hover:bg-gray-100'
          }`}
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            {fullView
              ? <><polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/><line x1="14" y1="10" x2="21" y2="3"/><line x1="3" y1="21" x2="10" y2="14"/></>
              : <><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></>
            }
          </svg>
          {fullView ? 'Compact' : 'Full View'}
        </button>
      </div>
      <div className={`overflow-auto scrollbar-thin ${fullView ? '' : 'max-h-[calc(100vh-340px)]'}`}>
        <table className="w-full text-[11px]">
          <thead className="sticky top-0">
            <tr className="bg-gray-700 text-white">
              <th className="px-2 py-2 text-center w-8">#</th>
              <th className="px-2 py-2 text-left">Cust Group</th>
              <th className="px-2 py-2 text-left">Customer</th>
              <th className="px-2 py-2 text-left">Origin</th>
              <th className="px-2 py-2 text-left">Item Code</th>
              <th className="px-2 py-2 text-left">Item Name</th>
              <th className="px-2 py-2 text-right cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('projKg')}>
                Proj KGs{sIcon('projKg')}
              </th>
              <th className="px-2 py-2 text-right cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('soKg')}>
                SO KGs{sIcon('soKg')}
              </th>
              <th className="px-2 py-2 text-right">Exp KGs</th>
              <th className="px-2 py-2 text-center cursor-pointer hover:bg-gray-600 select-none whitespace-nowrap" onClick={() => toggleSort('achKg')}>
                MTD Ach%{sIcon('achKg')}
              </th>
              <th className="px-2 py-2 text-center">SO vs Proj%</th>
              <th className="px-2 py-2 text-right">Diff KGs</th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((r, i) => (
              <tr key={`${r.itemCode}-${r.customer}-${i}`} className="border-b hover:bg-blue-50/40 even:bg-gray-50/50">
                <td className="px-2 py-1.5 text-center font-bold text-gray-400">{i + 1}</td>
                <td className="px-2 py-1.5">{r.custGroup}</td>
                <td className="px-2 py-1.5 max-w-[140px] truncate">{r.customer}</td>
                <td className="px-2 py-1.5">{r.origin}</td>
                <td className="px-2 py-1.5 font-medium">{r.itemCode}</td>
                <td className="px-2 py-1.5 max-w-[180px] truncate">{r.itemName}</td>
                <td className="px-2 py-1.5 text-right">{fmt(r.projKg)}</td>
                <td className="px-2 py-1.5 text-right">{fmt(r.soKg)}</td>
                <td className="px-2 py-1.5 text-right">{fmt(r.expKg)}</td>
                <td className="px-2 py-1.5 text-center"><span className={`px-2 py-0.5 rounded ${achClass(r.achKg)}`}>{fmtPct(r.achKg)}</span></td>
                <td className="px-2 py-1.5 text-center">{r.projKg > 0 ? fmtPct(r.soVsProj) : '-'}</td>
                <td className={`px-2 py-1.5 text-right ${r.diffKg < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(r.diffKg)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
