import { config } from '@/lib/config';
import type { SpagMatch, StandardsRuleKey } from '@/lib/types/ai-check';

// ── SPAG ─────────────────────────────────────────────────────────────────────

/**
 * Calls the SPAG webhook and returns filtered matches.
 * ignoredWords is pre-fetched by the caller (avoids session dependency).
 */
export async function callSpagWebhook(
  text: string,
  ignoredWords: Set<string>
): Promise<{ passed: boolean; errors: SpagMatch[] }> {
  type LTMatch = {
    offset: number;
    length: number;
    message: string;
    replacements: { value: string }[];
  };

  const response = await fetch(config.SPAG_WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ comment: text }),
  });

  if (!response.ok) {
    throw new Error(`SPAG service returned ${response.status}`);
  }

  let raw: unknown;
  try {
    raw = await response.json();
  } catch {
    throw new Error(`SPAG service returned non-JSON body (status ${response.status})`);
  }
  const unwrapped = Array.isArray(raw) ? (raw as unknown[])[0] : raw;
  const data = unwrapped as Record<string, unknown>;

  if (!Array.isArray(data?.matches)) {
    throw new Error(`Unexpected SPAG response shape: ${JSON.stringify(raw).slice(0, 200)}`);
  }

  const rawMatches = data.matches as LTMatch[];
  const errors: SpagMatch[] = rawMatches
    .map(m => ({
      word: text.slice(m.offset, m.offset + m.length),
      offset: m.offset,
      length: m.length,
      replacements: m.replacements.map(r => r.value).slice(0, 8),
      message: m.message,
    }))
    .filter(m => !ignoredWords.has(m.word.toLowerCase()));

  return { passed: errors.length === 0, errors };
}

// ── Standards ─────────────────────────────────────────────────────────────────

const STANDARDS_RULE_KEYS: StandardsRuleKey[] = [
  'UKSpelling', 'CourseOverviewIncluded', 'AcademicPerformanceIncluded',
  'TargetWordCountMet', 'TerminologyCorrect', 'JargonFree', 'DataSpecific',
  'ToneBalanced', 'SocialSkillsIncluded', 'CollaborationIncluded',
  'BehaviourIncluded', 'ParentalSupportIncluded', 'Formatting',
];

/**
 * Calls the Standards webhook and returns pass/fail with failed rule names.
 */
export async function callStandardsWebhook(
  text: string
): Promise<{ passed: boolean; failures: StandardsRuleKey[] }> {
  const response = await fetch(config.STANDARDS_WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ comment: text }),
  });

  if (!response.ok) {
    throw new Error(`Standards service returned ${response.status}`);
  }

  let raw: unknown;
  try {
    raw = await response.json();
  } catch {
    throw new Error(`Standards service returned non-JSON body (status ${response.status})`);
  }
  const unwrapped = Array.isArray(raw) ? (raw as unknown[])[0] : raw;
  const data = ((unwrapped as Record<string, unknown>)?.output ?? unwrapped) as Record<string, unknown>;

  // Simplified variant of the original toEntry() — intentionally drops instances/wordCount,
  // as the audit only needs pass/fail per rule.
  const toBool = (v: unknown): boolean => {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'string') return v.toLowerCase() === 'true' || v.toLowerCase() === 'passed';
    if (v && typeof v === 'object') return toBool((v as Record<string, unknown>).result);
    return false;
  };

  const failures: StandardsRuleKey[] = STANDARDS_RULE_KEYS.filter(
    key => !toBool(data?.[key])
  );

  return { passed: failures.length === 0, failures };
}
