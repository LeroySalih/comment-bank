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

/**
 * Calls the Standards webhook and returns pass/fail with failed rule names and per-rule instances.
 *
 * Service response key → StandardsRuleKey mapping:
 *   word_count          → TargetWordCountMet    {status, wordCount}
 *   terminology         → TerminologyCorrect    {status}
 *   HasDoubleNewlineslines → Formatting         {status}
 *   jargon              → JargonFree            {result, instances}
 *   negativity          → ToneBalanced          {result, instances}
 *   courseOverview      → CourseOverviewIncluded        plain string
 *   socialSkills        → SocialSkillsIncluded          plain string
 *   collaboration       → CollaborationIncluded         plain string
 *   behaviour           → BehaviourIncluded             plain string
 *   parentalSupport     → ParentalSupportIncluded       plain string
 *
 * UKSpelling and DataSpecific are not returned by the service and are excluded.
 */
export async function callStandardsWebhook(
  text: string
): Promise<{
  passed: boolean;
  failures: StandardsRuleKey[];
  failureDetails: Partial<Record<StandardsRuleKey, { instances: string[] }>>;
}> {
  const response = await fetch(config.STANDARDS_WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ comment: text }),
  });

  if (!response.ok) {
    throw new Error(`Standards service returned ${response.status}`);
  }

  const bodyText = await response.text();
  console.log('[standards-webhook] Raw response body:', bodyText);

  let raw: unknown;
  try {
    raw = JSON.parse(bodyText);
  } catch {
    console.error('[standards-webhook] Non-JSON response body:', bodyText);
    throw new Error(`Standards service returned non-JSON body (status ${response.status}): ${bodyText.slice(0, 300)}`);
  }

  const unwrapped = Array.isArray(raw) ? (raw as unknown[])[0] : raw;
  const d = (unwrapped ?? {}) as Record<string, unknown>;

  // Resolves plain strings, {status:...} and {result:...} objects to boolean
  const passed = (v: unknown): boolean => {
    if (typeof v === 'string') return v.toLowerCase() === 'passed';
    if (!v || typeof v !== 'object') return false;
    const o = v as Record<string, unknown>;
    const s = o.status ?? o.result ?? '';
    return typeof s === 'boolean' ? s : String(s).toLowerCase() === 'passed';
  };

  const getInstances = (v: unknown): string[] | undefined => {
    if (!v || typeof v !== 'object') return undefined;
    const arr = (v as Record<string, unknown>).instances;
    if (!Array.isArray(arr)) return undefined;
    const filtered = (arr as unknown[]).filter((x): x is string => typeof x === 'string' && x.length > 0);
    return filtered.length > 0 ? filtered : undefined;
  };

  // Explicit mapping from service keys to rule results
  const checks: [StandardsRuleKey, boolean, string[] | undefined][] = [
    ['TargetWordCountMet',        passed(d.wordCount),            undefined],
    ['TerminologyCorrect',        passed(d.terminology),          undefined],
    ['Formatting',                passed(d.hasDoubleNewlines),    undefined],
    ['JargonFree',                passed(d.jargon),               getInstances(d.jargon)],
    ['ToneBalanced',              passed(d.negativity),           getInstances(d.negativity)],
    ['CourseOverviewIncluded',    passed(d.courseOverview),       undefined],
    ['SocialSkillsIncluded',      passed(d.socialSkills),         undefined],
    ['CollaborationIncluded',     passed(d.collaboration),        undefined],
    ['BehaviourIncluded',         passed(d.behaviour),            undefined],
    ['ParentalSupportIncluded',   passed(d.parentalSupport),      undefined],
  ];

  const failures: StandardsRuleKey[] = [];
  const failureDetails: Partial<Record<StandardsRuleKey, { instances: string[] }>> = {};

  for (const [key, ok, instances] of checks) {
    if (!ok) {
      failures.push(key);
      if (instances && instances.length > 0) {
        failureDetails[key] = { instances };
      }
    }
  }

  return { passed: failures.length === 0, failures, failureDetails };
}
