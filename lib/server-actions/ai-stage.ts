'use server';

import { pool } from '@/lib/db';

export type AiStage = 'spag' | 'tone' | null;

export async function updateAiStage(
  assignmentId: string,
  stage: AiStage
): Promise<{ success: boolean }> {
  try {
    await pool.query(
      `UPDATE "Assignment" SET "aiStage" = $1 WHERE id = $2`,
      [stage, assignmentId]
    );
    return { success: true };
  } catch (error) {
    console.error('Failed to update aiStage:', error);
    return { success: false };
  }
}
