'use client';
import { useState } from 'react';
import { DashData } from '@/lib/helpers';
import UploadPage from '@/components/UploadPage';
import Dashboard from '@/components/Dashboard';

export default function Home() {
  const [data, setData] = useState<DashData | null>(null);

  if (!data) {
    return <UploadPage onDataReady={setData} />;
  }

  return <Dashboard data={data} onReUpload={() => setData(null)} />;
}
