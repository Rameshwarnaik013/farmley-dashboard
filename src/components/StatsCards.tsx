'use client';
import { Row, fmt, achClass, fmtPct } from '@/lib/helpers';
import { useMemo } from 'react';

export default function StatsCards({ rows }: { rows: Row[] }) {
  const stats = useMemo(() => {
    const total = rows.length;
    const darkRed = rows.filter(r => r.achKg > 120).length;
    const red = rows.filter(r => r.achKg > 100 && r.achKg <= 120).length;
    const orange = rows.filter(r => r.achKg >= 80 && r.achKg <= 100).length;
    const green = rows.filter(r => r.achKg > 0 && r.achKg < 80).length;
    const zero = rows.filter(r => r.achKg === 0 && r.projKg > 0).length;
    const totalSoKg = rows.reduce((s, r) => s + r.soKg, 0);
    const totalExpKg = rows.reduce((s, r) => s + r.expKg, 0);
    const ach = totalExpKg > 0 ? (totalSoKg / totalExpKg) * 100 : 0;
    return { total, darkRed, red, orange, green, zero, totalSoKg, ach };
  }, [rows]);

  const cards = [
    { label: 'Total SKUs', value: stats.total, color: 'text-brand-900' },
    { label: '>120% Urgent', value: stats.darkRed, color: 'text-[#c00000]' },
    { label: '100-120% Exceeding', value: stats.red, color: 'text-[#ff0000]' },
    { label: '80-100% Watch', value: stats.orange, color: 'text-[#e68a00]' },
    { label: '<80% Comfortable', value: stats.green, color: 'text-[#00b050]' },
    { label: 'Zero SO', value: stats.zero, color: 'text-gray-400' },
    { label: 'Total SO KGs', value: fmt(stats.totalSoKg, 0), color: 'text-brand-900' },
  ];

  return (
    <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3 mb-4">
      {cards.map(c => (
        <div key={c.label} className="bg-white rounded-lg shadow-sm p-3 text-center">
          <div className={`text-xl font-bold ${c.color}`}>{c.value}</div>
          <div className="text-[10px] text-gray-500 uppercase mt-1">{c.label}</div>
        </div>
      ))}
      <div className="bg-white rounded-lg shadow-sm p-3 text-center">
        <div className={`text-xl font-bold px-2 py-1 rounded ${achClass(stats.ach)} inline-block`}>
          {fmtPct(stats.ach)}
        </div>
        <div className="text-[10px] text-gray-500 uppercase mt-1">Overall Ach%</div>
      </div>
    </div>
  );
}
