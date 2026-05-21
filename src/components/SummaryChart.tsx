'use client';
import dynamic from 'next/dynamic';
import { Summary, achClass, fmtPct, fmt } from '@/lib/helpers';

const BarChart = dynamic(() => import('recharts').then(m => m.BarChart), { ssr: false, loading: () => <div className="h-64 bg-gray-100 animate-pulse rounded" /> });
const Bar = dynamic(() => import('recharts').then(m => m.Bar), { ssr: false });
const XAxis = dynamic(() => import('recharts').then(m => m.XAxis), { ssr: false });
const YAxis = dynamic(() => import('recharts').then(m => m.YAxis), { ssr: false });
const Tooltip = dynamic(() => import('recharts').then(m => m.Tooltip), { ssr: false });
const ResponsiveContainer = dynamic(() => import('recharts').then(m => m.ResponsiveContainer), { ssr: false });
const Cell = dynamic(() => import('recharts').then(m => m.Cell), { ssr: false });

function barColor(ach: number): string {
  if (ach > 120) return '#c00000';
  if (ach > 100) return '#ff0000';
  if (ach >= 80) return '#ffc000';
  if (ach >= 50) return '#92d050';
  return '#00b050';
}

export default function SummaryChart({ data, title }: { data: Summary[]; title: string }) {
  const totalExp = data.reduce((s, d) => s + d.expKg, 0);
  const totalSo = data.reduce((s, d) => s + d.soKg, 0);
  const overallAch = totalExp > 0 ? (totalSo / totalExp) * 100 : 0;

  return (
    <div>
      <div className="flex items-center gap-4 mb-4">
        <h3 className="text-lg font-bold text-brand-900">{title}</h3>
        <span className={`px-3 py-1 rounded text-sm ${achClass(overallAch)}`}>Overall: {fmtPct(overallAch)}</span>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4">
        <div className="bg-white rounded-lg shadow-sm p-4">
          <h4 className="text-sm font-semibold mb-3">Achievement % KGs</h4>
          <ResponsiveContainer width="100%" height={Math.max(200, data.length * 28)}>
            <BarChart data={data} layout="vertical" margin={{ left: 120, right: 40 }}>
              <XAxis type="number" domain={[0, 'auto']} tickFormatter={(v: number) => v + '%'} />
              <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 10 }} />
              <Tooltip formatter={(v: number) => fmtPct(v)} />
              <Bar dataKey="achKg" radius={[0, 4, 4, 0]}>
                {data.map((d, i) => <Cell key={i} fill={barColor(d.achKg)} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white rounded-lg shadow-sm overflow-auto max-h-96">
          <table className="w-full text-[11px]">
            <thead className="sticky top-0">
              <tr className="bg-gray-700 text-white">
                <th className="px-3 py-2 text-left">Name</th>
                <th className="px-2 py-2 text-right">#</th>
                <th className="px-2 py-2 text-right">Proj KGs</th>
                <th className="px-2 py-2 text-right">Exp KGs</th>
                <th className="px-2 py-2 text-right">SO KGs</th>
                <th className="px-2 py-2 text-center">Ach%</th>
                <th className="px-2 py-2 text-right">Diff</th>
              </tr>
            </thead>
            <tbody>
              {data.map(d => (
                <tr key={d.name} className="border-b hover:bg-blue-50/40">
                  <td className="px-3 py-1.5 font-medium">{d.name}</td>
                  <td className="px-2 py-1.5 text-right text-gray-400">{d.count}</td>
                  <td className="px-2 py-1.5 text-right">{fmt(d.projKg)}</td>
                  <td className="px-2 py-1.5 text-right">{fmt(d.expKg)}</td>
                  <td className="px-2 py-1.5 text-right">{fmt(d.soKg)}</td>
                  <td className="px-2 py-1.5 text-center"><span className={`px-2 py-0.5 rounded ${achClass(d.achKg)}`}>{fmtPct(d.achKg)}</span></td>
                  <td className={`px-2 py-1.5 text-right ${d.diffKg < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}`}>{fmt(d.diffKg)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
