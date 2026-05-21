'use client';
import { useMemo, useState } from 'react';
import { Row, achClass, fmt, fmtPct } from '@/lib/helpers';
import {
  useReactTable, getCoreRowModel, getSortedRowModel, getPaginationRowModel,
  flexRender, createColumnHelper, SortingState,
} from '@tanstack/react-table';

const ch = createColumnHelper<Row>();

const columns = [
  ch.accessor('custGroup', { header: 'Cust Group', size: 110 }),
  ch.accessor('customer', { header: 'Customer', size: 140 }),
  ch.accessor('origin', { header: 'Origin', size: 70 }),
  ch.accessor('itemCode', { header: 'Item Code', size: 130 }),
  ch.accessor('itemName', { header: 'Item Name', size: 200 }),
  ch.accessor('itemGroup', { header: 'Item Group', size: 130 }),
  ch.accessor('newMIS', { header: 'New MIS', size: 110 }),
  ch.accessor('projKg', { header: 'Proj KGs', size: 90, cell: i => fmt(i.getValue()) }),
  ch.accessor('soKg', { header: 'SO KGs', size: 90, cell: i => fmt(i.getValue()) }),
  ch.accessor('projUnits', { header: 'Proj Units', size: 90, cell: i => fmt(i.getValue(), 0) }),
  ch.accessor('expUnits', { header: 'Exp SO Units', size: 95, cell: i => fmt(i.getValue(), 0) }),
  ch.accessor('expKg', { header: 'Exp KGs', size: 90, cell: i => fmt(i.getValue()) }),
  ch.accessor('soUnits', { header: 'SO Units', size: 90, cell: i => fmt(i.getValue(), 0) }),
  ch.accessor('achKg', {
    header: 'MTD Ach%', size: 85,
    cell: i => <span className={`px-2 py-0.5 rounded text-xs ${achClass(i.getValue())}`}>{fmtPct(i.getValue())}</span>,
  }),
  ch.accessor('soVsProj', {
    header: 'SO vs Proj%', size: 90,
    cell: i => {
      const v = i.getValue();
      const projKg = i.row.original.projKg;
      return projKg > 0 ? fmtPct(v) : '-';
    },
  }),
  ch.accessor('soLeftPct', {
    header: '% SO Left', size: 85,
    cell: i => {
      const projKg = i.row.original.projKg;
      return projKg > 0 ? fmtPct(i.getValue()) : '-';
    },
  }),
  ch.accessor('soPctPerDay', {
    header: 'SO%/Day', size: 80,
    cell: i => i.getValue() > 0 ? fmtPct(i.getValue()) : '-',
  }),
  ch.accessor('daysToCover', {
    header: 'Days Left', size: 80,
    cell: i => i.getValue() > 0 ? fmt(i.getValue(), 1) : '-',
  }),
  ch.accessor('diffKg', {
    header: 'Diff KGs', size: 90,
    cell: i => <span className={i.getValue() < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}>{fmt(i.getValue())}</span>,
  }),
  ch.accessor('diffUnits', {
    header: 'Diff Units', size: 90,
    cell: i => <span className={i.getValue() < 0 ? 'text-red-700 font-semibold' : 'text-green-700'}>{fmt(i.getValue(), 0)}</span>,
  }),
  ch.accessor('runRateKg', { header: 'RunRate KGs', size: 95, cell: i => fmt(i.getValue()) }),
  ch.accessor('runRateUnits', { header: 'RunRate Units', size: 95, cell: i => fmt(i.getValue(), 0) }),
];

export default function DataTable({ data, pageSize = 50 }: { data: Row[]; pageSize?: number }) {
  const [sorting, setSorting] = useState<SortingState>([]);

  const table = useReactTable({
    data,
    columns,
    state: { sorting },
    onSortingChange: setSorting,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
    initialState: { pagination: { pageSize } },
  });

  return (
    <div>
      <div className="overflow-auto max-h-[calc(100vh-340px)] scrollbar-thin rounded-lg shadow-sm bg-white">
        <table className="w-full text-[11px]">
          <thead className="sticky top-0 z-10">
            {table.getHeaderGroups().map(hg => (
              <tr key={hg.id}>
                {hg.headers.map(h => (
                  <th
                    key={h.id}
                    onClick={h.column.getToggleSortingHandler()}
                    className="bg-gray-700 text-white px-2 py-2 text-left font-semibold cursor-pointer hover:bg-gray-600 whitespace-nowrap select-none"
                    style={{ minWidth: h.column.getSize() }}
                  >
                    {flexRender(h.column.columnDef.header, h.getContext())}
                    {{ asc: ' ▲', desc: ' ▼' }[h.column.getIsSorted() as string] ?? ''}
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {table.getRowModel().rows.map(row => (
              <tr key={row.id} className="border-b border-gray-100 hover:bg-blue-50/40 even:bg-gray-50/50">
                {row.getVisibleCells().map(cell => (
                  <td key={cell.id} className="px-2 py-1.5 whitespace-nowrap">
                    {flexRender(cell.column.columnDef.cell, cell.getContext())}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="flex items-center justify-between mt-3 text-xs text-gray-500">
        <span>{data.length} rows | Page {table.getState().pagination.pageIndex + 1} of {table.getPageCount()}</span>
        <div className="flex gap-1">
          <button onClick={() => table.setPageIndex(0)} disabled={!table.getCanPreviousPage()} className="px-2 py-1 border rounded disabled:opacity-30">{'<<'}</button>
          <button onClick={() => table.previousPage()} disabled={!table.getCanPreviousPage()} className="px-2 py-1 border rounded disabled:opacity-30">{'<'}</button>
          <button onClick={() => table.nextPage()} disabled={!table.getCanNextPage()} className="px-2 py-1 border rounded disabled:opacity-30">{'>'}</button>
          <button onClick={() => table.setPageIndex(table.getPageCount() - 1)} disabled={!table.getCanNextPage()} className="px-2 py-1 border rounded disabled:opacity-30">{'>>'}</button>
        </div>
      </div>
    </div>
  );
}
