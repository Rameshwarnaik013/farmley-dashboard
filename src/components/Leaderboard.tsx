'use client';
import { Row, achClass, fmt, fmtPct } from '@/lib/helpers';

export default function Leaderboard({ data, title, subtitle }: { data: Row[]; title: string; subtitle: string }) {
  return (
    <div className="bg-white rounded-lg shadow-sm overflow-auto">
      <div className="px-4 py-3 border-b bg-gray-50">
        <h3 className="text-sm font-bold text-brand-900">{title}</h3>
        <p className="text-[10px] text-gray-500">{subtitle}</p>
      </div>
      <table className="w-full text-[11px]">
        <thead className="sticky top-0">
          <tr className="bg-gray-700 text-white">
            <th className="px-2 py-2 text-center w-8">#</th>
            <th className="px-2 py-2 text-left">Cust Group</th>
            <th className="px-2 py-2 text-left">Customer</th>
            <th className="px-2 py-2 text-left">Origin</th>
            <th className="px-2 py-2 text-left">Item Code</th>
            <th className="px-2 py-2 text-left">Item Name</th>
            <th className="px-2 py-2 text-right">Proj KGs</th>
            <th className="px-2 py-2 text-right">Exp KGs</th>
            <th className="px-2 py-2 text-right">SO KGs</th>
            <th className="px-2 py-2 text-center">Ach% KGs</th>
            <th className="px-2 py-2 text-center">Ach% Units</th>
            <th className="px-2 py-2 text-right">Diff KGs</th>
            <th className="px-2 py-2 text-right">RunRate KGs</th>
          </tr>
        </thead>
        <tbody>
          {data.map((r, i) => (
            <tr key={`${r.itemCode}-${r.customer}-${i}`} className="border-b hover:bg-blue-50/40 even:bg-gray-50/50">
              <td className="px-2 py-1.5 text-center font-bold text-gray-400">{i + 1}</td>
              <td className="px-2 py-1.5">{r.custGroup}</td>
              <td className="px-2 py-1.5 max-w-[140px] truncate">{r.customer}</td>
              <td className="px-2 py-1.5">{r.origin}</td>
              <td className="px-2 py-1.5 font-medium">{r.itemCode}</td>
              <td className="px-2 py-1.5 max-w-[180px] truncate">{r.itemName}</td>
              <td className="px-2 py-1.5 text-right">{fmt(r.projKg)}</td>
              <td className="px-2 py-1.5 text-right">{fmt(r.expKg)}</td>
              <td className="px-2 py-1.5 text-right">{fmt(r.soKg)}</td>
              <td className="px-2 py-1.5 text-center"><span className={`px-2 py-0.5 rounded ${achClass(r.achKg)}`}>{fmtPct(r.achKg)}</span></td>
              <td className="px-2 py-1.5 text-center"><span className={`px-2 py-0.5 rounded ${achClass(r.achUnits)}`}>{fmtPct(r.achUnits)}</span></td>
              <td className={`px-2 py-1.5 text-right ${r.diffKg < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(r.diffKg)}</td>
              <td className="px-2 py-1.5 text-right">{fmt(r.runRateKg)}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
