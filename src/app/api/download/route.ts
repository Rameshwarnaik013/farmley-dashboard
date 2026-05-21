import { NextResponse } from 'next/server';
import { readFileSync, existsSync } from 'fs';
import { join } from 'path';

export async function GET() {
  // Try the generated Excel file
  const filePath = join(process.env.USERPROFILE || 'C:\\Users\\Admin', 'Downloads', 'RunRate_Final_Dashboard.xlsx');

  if (!existsSync(filePath)) {
    return NextResponse.json({ error: 'Excel file not found. Run the generation script first.' }, { status: 404 });
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
