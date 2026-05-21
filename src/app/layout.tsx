import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Farmley Run Rate Dashboard',
  description: 'Sales Order vs Projection Run Rate Tracker',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body className="min-h-screen">{children}</body>
    </html>
  );
}
