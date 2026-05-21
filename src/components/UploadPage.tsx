'use client';
import { useState, useRef } from 'react';
import { DashData } from '@/lib/helpers';
import { processExcelFiles } from '@/lib/processData';

interface Props {
  onDataReady: (data: DashData) => void;
}

export default function UploadPage({ onDataReady }: Props) {
  const [projFile, setProjFile] = useState<File | null>(null);
  const [soFile, setSoFile] = useState<File | null>(null);
  const [daysInMonth, setDaysInMonth] = useState(30);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState('');
  const projRef = useRef<HTMLInputElement>(null);
  const soRef = useRef<HTMLInputElement>(null);

  const canProcess = projFile && soFile && daysInMonth > 0;

  async function handleProcess() {
    if (!projFile || !soFile) return;
    setProcessing(true);
    setError('');
    try {
      const [projBuf, soBuf] = await Promise.all([
        projFile.arrayBuffer(),
        soFile.arrayBuffer(),
      ]);
      const data = processExcelFiles(projBuf, soBuf, daysInMonth);
      onDataReady(data);
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to process files. Check file format.');
      setProcessing(false);
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-brand-50 via-white to-brand-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl border border-brand-100 max-w-lg w-full p-8">
        {/* Logo & Title */}
        <div className="text-center mb-8">
          <div className="w-16 h-16 bg-gradient-to-br from-brand-700 to-brand-900 rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-lg">
            <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><line x1="10" y1="9" x2="8" y2="9"/></svg>
          </div>
          <h1 className="text-2xl font-bold text-brand-900">Farmley Run Rate Dashboard</h1>
          <p className="text-gray-500 text-sm mt-1">Upload your Excel files to generate the dashboard</p>
        </div>

        {/* File 1: Projection */}
        <div className="mb-5">
          <label className="block text-sm font-semibold text-gray-700 mb-2">
            1. Sales Projection File
            <span className="text-red-500 ml-0.5">*</span>
          </label>
          <div
            onClick={() => projRef.current?.click()}
            className={`border-2 border-dashed rounded-xl p-4 text-center cursor-pointer transition-all ${
              projFile ? 'border-green-400 bg-green-50' : 'border-gray-300 hover:border-brand-400 hover:bg-brand-50/30'
            }`}
          >
            <input
              ref={projRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={e => { setProjFile(e.target.files?.[0] || null); setError(''); }}
            />
            {projFile ? (
              <div className="flex items-center justify-center gap-2">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
                <span className="text-sm font-medium text-green-700">{projFile.name}</span>
                <span className="text-xs text-gray-400">({(projFile.size / 1024).toFixed(0)} KB)</span>
              </div>
            ) : (
              <div>
                <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="1.5" className="mx-auto mb-1"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
                <p className="text-sm text-gray-500">Click to upload <b>Projection Excel</b></p>
                <p className="text-[11px] text-gray-400 mt-0.5">e.g., May 2026 Sales Projection.xlsx</p>
              </div>
            )}
          </div>
        </div>

        {/* File 2: Sales Order */}
        <div className="mb-5">
          <label className="block text-sm font-semibold text-gray-700 mb-2">
            2. Sales Order Report File
            <span className="text-red-500 ml-0.5">*</span>
          </label>
          <div
            onClick={() => soRef.current?.click()}
            className={`border-2 border-dashed rounded-xl p-4 text-center cursor-pointer transition-all ${
              soFile ? 'border-green-400 bg-green-50' : 'border-gray-300 hover:border-brand-400 hover:bg-brand-50/30'
            }`}
          >
            <input
              ref={soRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={e => { setSoFile(e.target.files?.[0] || null); setError(''); }}
            />
            {soFile ? (
              <div className="flex items-center justify-center gap-2">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#16a34a" strokeWidth="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
                <span className="text-sm font-medium text-green-700">{soFile.name}</span>
                <span className="text-xs text-gray-400">({(soFile.size / 1024).toFixed(0)} KB)</span>
              </div>
            ) : (
              <div>
                <svg xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="1.5" className="mx-auto mb-1"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
                <p className="text-sm text-gray-500">Click to upload <b>Sales Order Excel</b></p>
                <p className="text-[11px] text-gray-400 mt-0.5">e.g., Custom Sales Order Report.xlsx</p>
              </div>
            )}
          </div>
        </div>

        {/* Days in Month */}
        <div className="mb-6">
          <label className="block text-sm font-semibold text-gray-700 mb-2">
            3. Number of Days in Month
          </label>
          <div className="flex items-center gap-3">
            <input
              type="number"
              value={daysInMonth}
              onChange={e => setDaysInMonth(Number(e.target.value) || 30)}
              min={1}
              max={31}
              className="w-24 px-3 py-2.5 border-2 border-gray-200 rounded-xl text-center text-lg font-bold text-brand-900 outline-none focus:border-brand-500 transition-all"
            />
            <span className="text-sm text-gray-500">days (used for daily demand calculation)</span>
          </div>
        </div>

        {/* Error */}
        {error && (
          <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-xl text-sm text-red-700">
            {error}
          </div>
        )}

        {/* Process Button */}
        <button
          onClick={handleProcess}
          disabled={!canProcess || processing}
          className={`w-full py-3 rounded-xl font-bold text-sm transition-all ${
            canProcess && !processing
              ? 'bg-gradient-to-r from-brand-700 to-brand-900 text-white hover:shadow-lg hover:shadow-brand-500/25 active:scale-[0.98]'
              : 'bg-gray-200 text-gray-400 cursor-not-allowed'
          }`}
        >
          {processing ? (
            <span className="flex items-center justify-center gap-2">
              <span className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></span>
              Processing...
            </span>
          ) : (
            'Generate Dashboard'
          )}
        </button>

        <p className="text-[11px] text-gray-400 text-center mt-4">
          All processing happens in your browser. No data is uploaded to any server.
        </p>
      </div>
    </div>
  );
}
