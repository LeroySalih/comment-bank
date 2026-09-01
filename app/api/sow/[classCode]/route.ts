import { NextRequest, NextResponse } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';

type Params = { params: Promise<{ classCode: string }> };

// GET /api/sow/[classCode]?academicYear=2026/27
export async function GET(req: NextRequest, { params }: Params) {
  const session = await getServerSession(authOptions);
  if (!session) return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });

  const { classCode } = await params;
  const { searchParams } = new URL(req.url);
  const academicYear = searchParams.get('academicYear') ?? currentAcademicYear();

  // Resolve class
  const classRow = await pool.query<{ id: string; name: string; subjectId: string }>(
    `SELECT c.id, c.name, c."subjectId", s.title as "subjectTitle"
     FROM "Class" c
     JOIN "Subject" s ON s.id = c."subjectId"
     WHERE c.name = $1`,
    [classCode],
  );
  if (!classRow.rows.length) return NextResponse.json({ error: 'Class not found' }, { status: 404 });
  const cls = classRow.rows[0] as { id: string; name: string; subjectId: string; subjectTitle: string };

  // Half-terms for this academic year
  const halfTerms = await pool.query<{
    id: string; label: string; startDate: string; endDate: string;
  }>(
    `SELECT id, label, "startDate"::text, "endDate"::text
     FROM "HalfTerm"
     WHERE "academicYear" = $1
     ORDER BY label`,
    [academicYear],
  );

  // SowUnits with a flag indicating if they have lessons in this academic year
  const units = await pool.query<{
    id: string; halfTermId: string; title: string; comment: string | null;
    isManual: boolean; hasLessons: boolean;
  }>(
    `SELECT
       su.id,
       su."halfTermId",
       su.title,
       su.comment,
       su."isManual",
       EXISTS (
         SELECT 1 FROM "Lesson" l
         WHERE l."sowUnitId" = su.id
           AND l.date >= ht."startDate"
           AND l.date <= ht."endDate"
       ) AS "hasLessons"
     FROM "SowUnit" su
     JOIN "HalfTerm" ht ON ht.id = su."halfTermId"
     WHERE su."classId" = $1
       AND ht."academicYear" = $2
     ORDER BY su.title`,
    [cls.id, academicYear],
  );

  return NextResponse.json({
    cls: { id: cls.id, name: cls.name, subjectTitle: cls.subjectTitle },
    academicYear,
    halfTerms: halfTerms.rows,
    units: units.rows,
  });
}

function currentAcademicYear(): string {
  const now = new Date();
  const year = now.getFullYear();
  // Academic year starts in September
  return now.getMonth() >= 8
    ? `${year}/${String(year + 1).slice(2)}`
    : `${year - 1}/${String(year).slice(2)}`;
}
