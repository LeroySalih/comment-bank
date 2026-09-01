import { NextRequest, NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';
import { randomUUID } from 'crypto';

type Params = { params: Promise<{ classCode: string }> };

// POST /api/sow/[classCode]/units — create a manual unit
export async function POST(req: NextRequest, { params }: Params) {
  const session = await getServerSession(authOptions);
  if (!session) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const { classCode } = await params;
  const { halfTermId, title } = await req.json();

  if (!halfTermId || !title?.trim()) {
    return NextResponse.json({ error: 'halfTermId and title are required' }, { status: 400 });
  }

  const classRow = await pool.query<{ id: string }>(
    `SELECT id FROM "Class" WHERE name = $1`,
    [classCode],
  );
  if (!classRow.rows.length) return NextResponse.json({ error: 'Class not found' }, { status: 404 });
  const classId = classRow.rows[0].id;

  // Upsert — merge with auto-created unit if one already exists for same title+halfTerm
  const result = await pool.query<{ id: string; title: string; comment: string | null; isManual: boolean }>(
    `INSERT INTO "SowUnit" (id, "classId", "halfTermId", title, "isManual")
     VALUES ($1, $2, $3, $4, true)
     ON CONFLICT ("classId", "halfTermId", title)
     DO UPDATE SET "isManual" = true
     RETURNING id, title, comment, "isManual"`,
    [randomUUID(), classId, halfTermId, title.trim()],
  );

  return NextResponse.json(result.rows[0], { status: 201 });
}
