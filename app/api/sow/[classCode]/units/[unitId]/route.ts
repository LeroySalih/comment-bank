import { NextRequest, NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';

type Params = { params: Promise<{ classCode: string; unitId: string }> };

// PATCH /api/sow/[classCode]/units/[unitId] — update comment
export async function PATCH(req: NextRequest, { params }: Params) {
  const session = await getServerSession(authOptions);
  if (!session) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const { unitId } = await params;
  const { comment } = await req.json();

  const result = await pool.query<{ id: string; comment: string | null }>(
    `UPDATE "SowUnit" SET comment = $1 WHERE id = $2 RETURNING id, comment`,
    [comment ?? null, unitId],
  );
  if (!result.rows.length) return NextResponse.json({ error: 'Not found' }, { status: 404 });

  return NextResponse.json(result.rows[0]);
}

// DELETE /api/sow/[classCode]/units/[unitId] — remove a manual unit (only if no lessons)
export async function DELETE(_req: NextRequest, { params }: Params) {
  const session = await getServerSession(authOptions);
  if (!session) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const { unitId } = await params;

  // Block deletion if lessons are assigned
  const lessonCheck = await pool.query<{ count: string }>(
    `SELECT COUNT(*) as count FROM "Lesson" WHERE "sowUnitId" = $1`,
    [unitId],
  );
  if (parseInt(lessonCheck.rows[0].count) > 0) {
    return NextResponse.json(
      { error: 'Unit has lessons assigned. Unassign lessons first.' },
      { status: 409 },
    );
  }

  await pool.query(`DELETE FROM "SowUnit" WHERE id = $1`, [unitId]);
  return new NextResponse(null, { status: 204 });
}
