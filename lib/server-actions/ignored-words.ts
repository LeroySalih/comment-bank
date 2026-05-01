'use server';

import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';

export async function addIgnoredWord(
  word: string
): Promise<{ success: boolean; error?: string }> {
  const session = await getServerSession(authOptions);
  if (!session?.user?.id) return { success: false, error: 'Not authenticated' };

  await pool.query(
    `INSERT INTO "IgnoredWord" (id, "teacherId", word)
     VALUES (gen_random_uuid()::text, $1, $2)
     ON CONFLICT ("teacherId", word) DO NOTHING`,
    [session.user.id, word.toLowerCase()]
  );

  return { success: true };
}

export async function getIgnoredWords(): Promise<string[]> {
  const session = await getServerSession(authOptions);
  if (!session?.user?.id) return [];

  const { rows } = await pool.query<{ word: string }>(
    `SELECT word FROM "IgnoredWord" WHERE "teacherId" = $1`,
    [session.user.id]
  );

  return rows.map(r => r.word);
}
