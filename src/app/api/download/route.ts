import { NextResponse } from 'next/server';
import { readFileSync, existsSync } from 'fs';
import { join } from 'path';

export async function GET() {
  // Try public folder first (works on Vercel), then local Downloads (dev fallback)
  const publicPath = join(process.cwd(), 'public', 'RunRate_Final_Dashboard.xlsx');
  const localPath = join(process.env.USERPROFILE || process.env.HOME || '', 'Downloads', 'RunRate_Final_Dashboard.xlsx');

  const filePath = existsSync(publicPath) ? publicPath : localPath;

  if (!existsSync(filePath)) {
    return NextResponse.json({ error: 'Excel file not found.' }, { status: 404 });
  }

  const fileBuffer = readFileSync(filePath);
  const filename = 'RunRate_Dashboard_' + new Date().toISOString().slice(0, 10) + '.xlsx';

  return new NextResponse(fileBuffer, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': fileBuffer.length.toString(),
    },
  });
}
