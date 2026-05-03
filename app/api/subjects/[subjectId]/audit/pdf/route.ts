export const runtime = 'nodejs';

import { NextRequest } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { consumePdf } from '@/lib/audit/pdf-store';

export async function GET(
  request: NextRequest,
  { params }: { params: Promise<{ subjectId: string }> }
) {
  const { subjectId } = await params;

  const session = await getServerSession(authOptions);
  if (!session?.user?.id) {
    return new Response('Unauthorized', { status: 401 });
  }

  const token = request.nextUrl.searchParams.get('token');
  if (!token) {
    return new Response('Missing token', { status: 400 });
  }

  const buffer = consumePdf(token);
  if (!buffer) {
    return new Response('PDF not found or expired', { status: 404 });
  }

  const filename = `audit-${subjectId}-${Date.now()}.pdf`;

  return new Response(buffer as any, {
    headers: {
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': String(buffer.length),
    },
  });
}
