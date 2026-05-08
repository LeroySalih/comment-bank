'use server';

import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';
import type { SpagMatch, SpagResult, StandardsResult, StandardsRuleEntry } from '@/lib/types/ai-check';
import { config } from '@/lib/config';

// ── SPAG CHECK ──────────────────────────────────────────────────────────────

export async function requestSpagCheck(
  commentText: string
): Promise<{ success: true; result: SpagResult } | { success: true; result: null } | { success: false; error: string }> {
  if (!commentText.trim()) {
    return { success: false, error: 'No comment text to check' };
  }

  const session = await getServerSession(authOptions);
  if (!session?.user?.id) {
    return { success: false, error: 'Not authenticated' };
  }

  const { rows: ignoredRows } = await pool.query<{ word: string }>(
    `SELECT word FROM "IgnoredWord" WHERE "teacherId" = $1`,
    [session.user.id]
  );
  const ignoredWords = new Set(ignoredRows.map(r => r.word.toLowerCase()));

  console.log('[spag-check] Calling SPAG service:', config.SPAG_WEBHOOK_URL);

  let response: Response;
  try {
    response = await fetch(config.SPAG_WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ comment: commentText }),
    });
  } catch (err) {
    console.error('[spag-check] Fetch failed for URL:', config.SPAG_WEBHOOK_URL, err);
    return { success: false, error: `Could not reach SPAG service (${config.SPAG_WEBHOOK_URL})` };
  }

  if (!response.ok) {
    console.error('[spag-check] Non-OK status', response.status, 'for URL:', config.SPAG_WEBHOOK_URL);
    return { success: false, error: `SPAG service returned ${response.status} (${config.SPAG_WEBHOOK_URL})` };
  }

  let raw: unknown;
  try {
    raw = await response.json();
  } catch (err) {
    console.error('[spag-check] JSON parse failed for URL:', config.SPAG_WEBHOOK_URL, err);
    return { success: false, error: `SPAG service returned an empty or invalid response (${config.SPAG_WEBHOOK_URL})` };
  }

  try {
    const unwrapped = Array.isArray(raw) ? (raw as unknown[])[0] : raw;
    const data = unwrapped as Record<string, unknown>;

    if (!Array.isArray(data?.matches)) {
      console.error('[spag-check] Unexpected response shape from URL:', config.SPAG_WEBHOOK_URL, '\nResponse:', JSON.stringify(raw, null, 2));
      return { success: false, error: `Unexpected response shape from ${config.SPAG_WEBHOOK_URL}: ${JSON.stringify(raw).slice(0, 200)}` };
    }

    type LTMatch = {
      offset: number;
      length: number;
      message: string;
      replacements: { value: string }[];
    };

    const rawMatches = data.matches as LTMatch[];

    const matches: SpagMatch[] = rawMatches
      .map(m => ({
        word: commentText.slice(m.offset, m.offset + m.length),
        offset: m.offset,
        length: m.length,
        replacements: m.replacements.map(r => r.value).slice(0, 8),
        message: m.message,
      }))
      .filter(m => !ignoredWords.has(m.word.toLowerCase()));

    console.log(`[spag-check] LanguageTool returned ${rawMatches.length} matches; ${matches.length} after ignoring teacher's ignored words`);

    if (matches.length === 0) {
      return { success: true, result: null };
    }

    return { success: true, result: { matches } };
  } catch (err) {
    console.error('[spag-check] Parse exception for URL:', config.SPAG_WEBHOOK_URL, err);
    return { success: false, error: `Failed to parse SPAG response from ${config.SPAG_WEBHOOK_URL}` };
  }
}

// ── STANDARDS CHECK ─────────────────────────────────────────────────────────

const STANDARDS_RULE_KEYS = [
  'UKSpelling',
  'CourseOverviewIncluded',
  'AcademicPerformanceIncluded',
  'TargetWordCountMet',
  'TerminologyCorrect',
  'JargonFree',
  'DataSpecific',
  'ToneBalanced',
  'SocialSkillsIncluded',
  'CollaborationIncluded',
  'BehaviourIncluded',
  'ParentalSupportIncluded',
  'Formatting',
] as const;

export async function requestStandardsCheck(
  commentText: string
): Promise<{ success: true; result: StandardsResult; raw: unknown } | { success: false; error: string }> {
  if (!commentText.trim()) {
    return { success: false, error: 'No comment text to check' };
  }

  console.log('[standards-check] Calling service:', config.STANDARDS_WEBHOOK_URL);

  let response: Response;
  try {
    response = await fetch(config.STANDARDS_WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ comment: commentText }),
    });
  } catch (err) {
    console.error('[standards-check] Fetch failed:', err);
    return { success: false, error: `Could not reach Standards service (${config.STANDARDS_WEBHOOK_URL})` };
  }

  if (!response.ok) {
    console.error('[standards-check] Non-OK status', response.status);
    return { success: false, error: `Standards service returned ${response.status} (${config.STANDARDS_WEBHOOK_URL})` };
  }

  let raw: unknown;
  try {
    raw = await response.json();
  } catch (err) {
    console.error('[standards-check] JSON parse failed:', err);
    return { success: false, error: `Standards service returned an empty or invalid response (${config.STANDARDS_WEBHOOK_URL})` };
  }

  try {
    const unwrapped = Array.isArray(raw) ? (raw as unknown[])[0] : raw;
    const obj = unwrapped as Record<string, unknown>;

    const statusPassed = (v: unknown): boolean => {
      if (typeof v === 'string') return v.toLowerCase() === 'passed';
      if (!v || typeof v !== 'object') return false;
      const o = v as Record<string, unknown>;
      const s = (o.status ?? o.result ?? '');
      return typeof s === 'boolean' ? s : String(s).toLowerCase() === 'passed';
    };

    const instances = (v: unknown): string[] | undefined => {
      if (!v || typeof v !== 'object') return undefined;
      const arr = (v as Record<string, unknown>).instances;
      if (!Array.isArray(arr)) return undefined;
      const filtered = (arr as unknown[]).filter((x): x is string => typeof x === 'string' && x.length > 0);
      return filtered.length > 0 ? filtered : undefined;
    };

    const wordCountObj = obj.word_count as Record<string, unknown> | undefined;

    const result: StandardsResult = {
      // Not returned by service — omit from failing checks by defaulting to true
      UKSpelling:   { result: true },
      DataSpecific: { result: true },

      TargetWordCountMet: {
        result: statusPassed(obj.word_count),
        wordCount: typeof wordCountObj?.wordCount === 'number' ? wordCountObj.wordCount : undefined,
      },

      TerminologyCorrect: { result: statusPassed(obj.terminology) },
      Formatting:         { result: statusPassed(obj.HasDoubleNewlineslines) },

      JargonFree:    { result: statusPassed(obj.jargon),     instances: instances(obj.jargon) },
      ToneBalanced:  { result: statusPassed(obj.negativity), instances: instances(obj.negativity) },

      CourseOverviewIncluded:      { result: statusPassed(obj.courseOverview) },
      AcademicPerformanceIncluded: { result: statusPassed(obj.academicPerformance) },
      SocialSkillsIncluded:        { result: statusPassed(obj.socialSkills) },
      CollaborationIncluded:       { result: statusPassed(obj.collaboration) },
      BehaviourIncluded:           { result: statusPassed(obj.behaviour) },
      ParentalSupportIncluded:     { result: statusPassed(obj.parentalSupport) },

      Status: { result: false }, // derived below
    };

    result.Status = { result: STANDARDS_RULE_KEYS.every(k => result[k].result === true) };

    return { success: true, result, raw };
  } catch (err) {
    console.error('[standards-check] Parse exception:', err);
    return { success: false, error: `Failed to parse Standards response from ${config.STANDARDS_WEBHOOK_URL}` };
  }
}
