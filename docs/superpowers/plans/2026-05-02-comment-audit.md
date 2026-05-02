# Comment Audit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a subject-level audit tool that SPAG-checks every comment individually, runs 50 structured sample reports through the standards checker, and produces a downloadable PDF — triggered from a modal on the subject admin page.

**Architecture:** A streaming `GET /api/subjects/[subjectId]/audit` route sends newline-delimited JSON events via `ReadableStream` (SSE). The client modal subscribes with `EventSource`, updates two progress bars, then downloads a PDF once complete. PDF generation uses `@react-pdf/renderer` server-side; the buffer is stored behind a short-lived UUID token.

**Tech Stack:** Next.js 16 App Router, TypeScript, `pg`, `@react-pdf/renderer`, Vitest, Tailwind CSS

---

## File Map

| File | Status | Responsibility |
|---|---|---|
| `lib/audit/types.ts` | Create | All audit-specific TypeScript types |
| `lib/audit/substitute-variables.ts` | Create | Replace `<Name>` etc. with fixed test values |
| `lib/audit/build-reports.ts` | Create | Assemble up to 50 sample report texts |
| `lib/audit/webhook-calls.ts` | Create | Pure webhook calls (no session dependency) |
| `lib/audit/pdf-store.ts` | Create | In-memory token → PDF buffer store with TTL |
| `lib/audit/generate-pdf.ts` | Create | `@react-pdf/renderer` document + render function |
| `app/api/subjects/[subjectId]/audit/route.ts` | Create | SSE streaming audit handler |
| `app/api/subjects/[subjectId]/audit/pdf/route.ts` | Create | Token-gated PDF download |
| `components/AuditModal.tsx` | Create | Modal UI with EventSource client + progress bars |
| `app/hod/subject/[subjectId]/page.tsx` | Modify | Add Audit button that opens AuditModal |
| `lib/__tests__/audit-substitute-variables.test.ts` | Create | Unit tests for variable substitution |
| `lib/__tests__/audit-build-reports.test.ts` | Create | Unit tests for the 50-report sampling logic |

---

## Task 1: Install `@react-pdf/renderer`

**Files:**
- Modify: `package.json`

- [ ] **Step 1: Install the package**

```bash
cd /Users/leroysalih/nodejs/comment-bank
npm install @react-pdf/renderer
```

Expected: `@react-pdf/renderer` appears in `package.json` dependencies. No peer-dep errors.

- [ ] **Step 2: Verify TypeScript types are included**

```bash
node -e "require('@react-pdf/renderer'); console.log('OK')"
```

Expected: prints `OK` with no errors.

- [ ] **Step 3: Commit**

```bash
git add package.json package-lock.json
git commit -m "chore: add @react-pdf/renderer for server-side PDF generation"
```

---

## Task 2: Audit Types

**Files:**
- Create: `lib/audit/types.ts`

- [ ] **Step 1: Create the types file**

```typescript
// lib/audit/types.ts

import type { SpagMatch, StandardsRuleKey } from '@/lib/types/ai-check';

// ── SSE event union ──────────────────────────────────────────────────────────

export type AuditInitEvent = {
  type: 'init';
  totalComments: number;
  totalReports: number;
};

export type AuditSpagEvent = {
  type: 'spag';
  code: string;
  groupName: string;
  passed: boolean;
  errors: SpagMatch[];
};

export type AuditSpagDoneEvent = {
  type: 'spag_done';
};

export type AuditStandardsEvent = {
  type: 'standards';
  reportIndex: number;
  /** groupId → selected comment code */
  codes: Record<string, string>;
  passed: boolean;
  failures: StandardsRuleKey[];
};

export type AuditStandardsDoneEvent = {
  type: 'standards_done';
};

export type AuditUntestedEvent = {
  type: 'untested';
  items: { code: string; groupName: string }[];
};

export type AuditCompleteEvent = {
  type: 'complete';
  pdfUrl: string;
};

export type AuditErrorEvent = {
  type: 'error';
  message: string;
};

export type AuditEvent =
  | AuditInitEvent
  | AuditSpagEvent
  | AuditSpagDoneEvent
  | AuditStandardsEvent
  | AuditStandardsDoneEvent
  | AuditUntestedEvent
  | AuditCompleteEvent
  | AuditErrorEvent;

// ── Data structures ──────────────────────────────────────────────────────────

/** One comment option with its group context, ready for substitution */
export type AuditCommentEntry = {
  code: string;
  text: string;
  groupId: string;
  groupName: string;
};

/** One sampled report — group selections + assembled text */
export type SampledReport = {
  reportIndex: number;
  /** groupId → { code, text, groupTitle } */
  selections: Record<string, { code: string; text: string; groupTitle: string }>;
  assembledText: string;
};

/** SPAG result accumulated during Phase 1 */
export type SpagAuditEntry = {
  code: string;
  groupName: string;
  rawText: string;
  passed: boolean;
  errors: SpagMatch[];
};

/** Standards result accumulated during Phase 2 */
export type StandardsAuditEntry = {
  reportIndex: number;
  codes: Record<string, string>;
  passed: boolean;
  failures: StandardsRuleKey[];
};

/** Full data passed to the PDF renderer */
export type AuditPdfData = {
  subjectTitle: string;
  subjectCode: string;
  generatedAt: Date;
  totalReports: number;
  passedReports: number;
  spagEntries: SpagAuditEntry[];
  standardsFailures: StandardsAuditEntry[];
  untestedItems: { code: string; groupName: string }[];
  /** All subject groups (non-linked), for the comments section */
  groupTitles: Record<string, string>;
};
```

- [ ] **Step 2: Run TypeScript check**

```bash
npx tsc --noEmit --incremental false 2>&1 | head -20
```

Expected: no errors related to `lib/audit/types.ts`.

- [ ] **Step 3: Commit**

```bash
git add lib/audit/types.ts
git commit -m "feat(audit): add audit TypeScript types"
```

---

## Task 3: Variable Substitution

**Files:**
- Create: `lib/audit/substitute-variables.ts`
- Create: `lib/__tests__/audit-substitute-variables.test.ts`

- [ ] **Step 1: Write the failing tests**

```typescript
// lib/__tests__/audit-substitute-variables.test.ts

import { describe, it, expect } from 'vitest';
import { substituteVariables } from '../audit/substitute-variables';

describe('substituteVariables', () => {
  it('replaces <Name> with Alex', () => {
    expect(substituteVariables('<Name> is a great student.', 'Computing')).toBe(
      'Alex is a great student.'
    );
  });

  it('replaces <he/she> with they', () => {
    expect(substituteVariables('<he/she> works hard.', 'Computing')).toBe(
      'they works hard.'
    );
  });

  it('replaces <his/her> with their', () => {
    expect(substituteVariables('<his/her> work is excellent.', 'Computing')).toBe(
      'their work is excellent.'
    );
  });

  it('replaces <him/her> with them', () => {
    expect(substituteVariables('I encourage <him/her>.', 'Computing')).toBe(
      'I encourage them.'
    );
  });

  it('replaces <Subject> with the provided subject title', () => {
    expect(substituteVariables('<Name> enjoys <Subject>.', 'French')).toBe(
      'Alex enjoys French.'
    );
  });

  it('replaces <Year> with Year 10', () => {
    expect(substituteVariables('In <Year>, <Name> excelled.', 'Maths')).toBe(
      'In Year 10, Alex excelled.'
    );
  });

  it('replaces <EoYLevel> with 6', () => {
    expect(substituteVariables('Achieved <EoYLevel>.', 'Maths')).toBe(
      'Achieved 6.'
    );
  });

  it('replaces <TargetLevel> with 7', () => {
    expect(substituteVariables('Target is <TargetLevel>.', 'Maths')).toBe(
      'Target is 7.'
    );
  });

  it('replaces all occurrences of the same variable', () => {
    expect(substituteVariables('<Name> and <Name> again.', 'Computing')).toBe(
      'Alex and Alex again.'
    );
  });

  it('returns text unchanged when no variables present', () => {
    expect(substituteVariables('No variables here.', 'Computing')).toBe(
      'No variables here.'
    );
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

```bash
npx vitest run lib/__tests__/audit-substitute-variables.test.ts 2>&1 | tail -10
```

Expected: FAIL — `Cannot find module '../audit/substitute-variables'`

- [ ] **Step 3: Implement substitute-variables**

```typescript
// lib/audit/substitute-variables.ts

const FIXED_VALUES: Record<string, string> = {
  '<Name>': 'Alex',
  '<he/she>': 'they',
  '<his/her>': 'their',
  '<him/her>': 'them',
  '<Year>': 'Year 10',
  '<EoYLevel>': '6',
  '<TargetLevel>': '7',
};

/**
 * Replaces all template variables in a comment with fixed audit test values.
 * <Subject> is replaced with the actual subject title from the DB.
 */
export function substituteVariables(text: string, subjectTitle: string): string {
  let result = text;
  for (const [variable, value] of Object.entries(FIXED_VALUES)) {
    result = result.replaceAll(variable, value);
  }
  result = result.replaceAll('<Subject>', subjectTitle);
  return result;
}
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
npx vitest run lib/__tests__/audit-substitute-variables.test.ts
```

Expected: all 10 tests pass.

- [ ] **Step 5: Commit**

```bash
git add lib/audit/substitute-variables.ts lib/__tests__/audit-substitute-variables.test.ts
git commit -m "feat(audit): add variable substitution with fixed test values"
```

---

## Task 4: Report Builder

**Files:**
- Create: `lib/audit/build-reports.ts`
- Create: `lib/__tests__/audit-build-reports.test.ts`

- [ ] **Step 1: Write the failing tests**

```typescript
// lib/__tests__/audit-build-reports.test.ts

import { describe, it, expect } from 'vitest';
import { buildSampleReports } from '../audit/build-reports';

type Option = { id: string; code: string; text: string; displayOrder: number };
type Group = { id: string; title: string; isLinked: boolean; options: Option[] };

function makeGroup(id: string, title: string, codes: string[], isLinked = false): Group {
  return {
    id,
    title,
    isLinked,
    options: codes.map((code, i) => ({
      id: `${id}-opt-${i}`,
      code,
      text: `${title} ${code} comment text`,
      displayOrder: i + 1,
    })),
  };
}

const COMMON: Group[] = [
  makeGroup('cg1', 'Effort', ['Excellent', 'Good', 'Poor']),
];

describe('buildSampleReports', () => {
  it('returns at most maxReports reports', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'Skills', ['High', 'Med', 'Low']),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 5);
    expect(reports.length).toBeLessThanOrEqual(5);
  });

  it('skips isLinked groups', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'LinkedGroup', ['A', 'B'], true),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 50);
    // Only g1 should appear in selections, not g2
    for (const r of reports) {
      expect(r.selections['g2']).toBeUndefined();
    }
  });

  it('includes common group options in assembled text', () => {
    const groups = [makeGroup('g1', 'Knowledge', ['High'])];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 50);
    expect(reports.length).toBeGreaterThan(0);
    // Common group text should be in the assembled output
    expect(reports[0].assembledText).toContain('Effort');
  });

  it('generates "all high" report as first report', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'Skills', ['High', 'Med', 'Low']),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 50);
    // First report: lowest displayOrder (index 0) from each group = "High"
    expect(reports[0].selections['g1'].code).toBe('High');
    expect(reports[0].selections['g2'].code).toBe('High');
  });

  it('tracks untested codes', () => {
    // 2 groups × 3 options each = 6 codes. With maxReports=1, some will be untested.
    const groups = [
      makeGroup('g1', 'Knowledge', ['H', 'M', 'L']),
      makeGroup('g2', 'Skills', ['H', 'M', 'L']),
    ];
    const { untestedItems } = buildSampleReports(groups, COMMON, 'Computing', 1);
    expect(untestedItems.length).toBeGreaterThan(0);
  });

  it('reports zero untested when all codes are covered', () => {
    // 1 group × 1 option = fully covered in 1 report
    const groups = [makeGroup('g1', 'Knowledge', ['OnlyOption'])];
    const { untestedItems } = buildSampleReports(groups, COMMON, 'Computing', 50);
    expect(untestedItems.length).toBe(0);
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
npx vitest run lib/__tests__/audit-build-reports.test.ts 2>&1 | tail -10
```

Expected: FAIL — `Cannot find module '../audit/build-reports'`

- [ ] **Step 3: Implement build-reports**

```typescript
// lib/audit/build-reports.ts

import { substituteVariables } from './substitute-variables';
import type { SampledReport } from './types';

type Option = { id: string; code: string; text: string; displayOrder: number };
type Group = { id: string; title: string; isLinked: boolean; options: Option[] };

type BuildResult = {
  reports: SampledReport[];
  untestedItems: { code: string; groupName: string }[];
};

/** Sort options by displayOrder ascending; index 0 = "High" */
function sortedOptions(options: Option[]): Option[] {
  return [...options].sort((a, b) => a.displayOrder - b.displayOrder);
}

function pickOption(options: Option[], tier: 'high' | 'medium' | 'low'): Option {
  const sorted = sortedOptions(options);
  if (tier === 'high') return sorted[0];
  if (tier === 'low') return sorted[sorted.length - 1];
  return sorted[Math.floor((sorted.length - 1) / 2)];
}

function assembleText(
  subjectSelections: Record<string, Option>,
  commonSelections: Record<string, Option>,
  subjectTitle: string
): string {
  const parts: string[] = [];
  for (const opt of Object.values(subjectSelections)) {
    parts.push(substituteVariables(opt.text, subjectTitle));
  }
  for (const opt of Object.values(commonSelections)) {
    parts.push(substituteVariables(opt.text, subjectTitle));
  }
  return parts.join(' ');
}

export function buildSampleReports(
  subjectGroups: Group[],
  commonGroups: Group[],
  subjectTitle: string,
  maxReports = 50
): BuildResult {
  // Filter out linked groups; groups must have at least one option
  const activeGroups = subjectGroups.filter(g => !g.isLinked && g.options.length > 0);

  // Fixed common selections: first option by displayOrder from each common group
  const commonSelections: Record<string, Option> = {};
  for (const cg of commonGroups) {
    if (!cg.isLinked && cg.options.length > 0) {
      commonSelections[cg.id] = sortedOptions(cg.options)[0];
    }
  }

  const reports: SampledReport[] = [];

  // Track which (groupId, code) combos have been seen
  const seenCodes = new Set<string>(); // `${groupId}:${code}`

  function addReport(subjectSelections: Record<string, Option>): boolean {
    if (reports.length >= maxReports) return false;
    const index = reports.length;
    const selectionsForReport: SampledReport['selections'] = {};
    for (const [groupId, opt] of Object.entries(subjectSelections)) {
      const group = activeGroups.find(g => g.id === groupId)!;
      selectionsForReport[groupId] = {
        code: opt.code,
        text: opt.text,
        groupTitle: group.title,
      };
      seenCodes.add(`${groupId}:${opt.code}`);
    }
    // Also record common selections
    for (const [groupId, opt] of Object.entries(commonSelections)) {
      const cg = commonGroups.find(g => g.id === groupId)!;
      selectionsForReport[groupId] = {
        code: opt.code,
        text: opt.text,
        groupTitle: cg.title,
      };
    }
    const assembledText = assembleText(subjectSelections, commonSelections, subjectTitle);
    reports.push({ reportIndex: index, selections: selectionsForReport, assembledText });
    return true;
  }

  // ── Strategy 1: All High ────────────────────────────────────────────────
  {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'high');
    addReport(sel);
  }

  // ── Strategy 2: All Medium ──────────────────────────────────────────────
  if (reports.length < maxReports) {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'medium');
    addReport(sel);
  }

  // ── Strategy 3: All Low ─────────────────────────────────────────────────
  if (reports.length < maxReports) {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'low');
    addReport(sel);
  }

  // ── Strategy 4: Mostly High (one group rotated to Medium) ───────────────
  for (const rotateGroup of activeGroups) {
    if (reports.length >= maxReports) break;
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) {
      sel[g.id] = g.id === rotateGroup.id
        ? pickOption(g.options, 'medium')
        : pickOption(g.options, 'high');
    }
    addReport(sel);
  }

  // ── Strategy 5: Mostly Low (one group rotated to High) ──────────────────
  for (const rotateGroup of activeGroups) {
    if (reports.length >= maxReports) break;
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) {
      sel[g.id] = g.id === rotateGroup.id
        ? pickOption(g.options, 'high')
        : pickOption(g.options, 'low');
    }
    addReport(sel);
  }

  // ── Strategy 6: Coverage fill ────────────────────────────────────────────
  // Collect all option codes not yet seen
  while (reports.length < maxReports) {
    const unseenOptions: { groupId: string; opt: Option }[] = [];
    for (const g of activeGroups) {
      for (const opt of g.options) {
        if (!seenCodes.has(`${g.id}:${opt.code}`)) {
          unseenOptions.push({ groupId: g.id, opt });
        }
      }
    }
    if (unseenOptions.length === 0) break;

    // Build one report that covers as many unseen options as possible
    const sel: Record<string, Option> = {};
    const coveredGroups = new Set<string>();

    // First pass: pick unseen options
    for (const { groupId, opt } of unseenOptions) {
      if (!coveredGroups.has(groupId)) {
        sel[groupId] = opt;
        coveredGroups.add(groupId);
      }
    }
    // Fill remaining groups with their "high" option
    for (const g of activeGroups) {
      if (!sel[g.id]) sel[g.id] = pickOption(g.options, 'high');
    }
    addReport(sel);
  }

  // ── Untested codes ───────────────────────────────────────────────────────
  const untestedItems: { code: string; groupName: string }[] = [];
  for (const g of activeGroups) {
    for (const opt of g.options) {
      if (!seenCodes.has(`${g.id}:${opt.code}`)) {
        untestedItems.push({ code: opt.code, groupName: g.title });
      }
    }
  }

  return { reports, untestedItems };
}
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
npx vitest run lib/__tests__/audit-build-reports.test.ts
```

Expected: all 6 tests pass.

- [ ] **Step 5: Commit**

```bash
git add lib/audit/build-reports.ts lib/__tests__/audit-build-reports.test.ts
git commit -m "feat(audit): add 50-report sampling logic with coverage tracking"
```

---

## Task 5: Webhook Calls (Pure Functions)

**Files:**
- Create: `lib/audit/webhook-calls.ts`

The existing `requestSpagCheck` in `lib/server-actions/ai-check.ts` is a server action — it uses `getServerSession` and cannot be called from an API route. This task extracts the core HTTP logic into pure functions that accept pre-fetched inputs.

- [ ] **Step 1: Create webhook-calls.ts**

```typescript
// lib/audit/webhook-calls.ts

import { config } from '@/lib/config';
import type { SpagMatch, StandardsRuleKey } from '@/lib/types/ai-check';
import type { StandardsRuleEntry } from '@/lib/types/ai-check';

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

  const raw: unknown = await response.json();
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

  const raw: unknown = await response.json();
  const unwrapped = Array.isArray(raw) ? (raw as unknown[])[0] : raw;
  const data = ((unwrapped as Record<string, unknown>)?.output ?? unwrapped) as Record<string, unknown>;

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
```

- [ ] **Step 2: Type-check**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "webhook-calls" | head -10
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add lib/audit/webhook-calls.ts
git commit -m "feat(audit): add pure webhook call functions for SPAG and standards"
```

---

## Task 6: PDF Store

**Files:**
- Create: `lib/audit/pdf-store.ts`

- [ ] **Step 1: Create pdf-store.ts**

```typescript
// lib/audit/pdf-store.ts

import { randomUUID } from 'crypto';

type Entry = { buffer: Buffer; expiresAt: number };

// Module-level Map — persists for the lifetime of the Node.js process.
// Both routes must use `export const runtime = 'nodejs'` to share this instance.
const store = new Map<string, Entry>();

const TTL_MS = 5 * 60 * 1000; // 5 minutes

/** Store a PDF buffer and return a short-lived token to retrieve it. */
export function storePdf(buffer: Buffer): string {
  const token = randomUUID();
  store.set(token, { buffer, expiresAt: Date.now() + TTL_MS });
  // Lazy cleanup: remove expired entries on each store
  for (const [key, entry] of store) {
    if (entry.expiresAt < Date.now()) store.delete(key);
  }
  return token;
}

/** Retrieve and delete a PDF buffer by token. Returns null if expired or not found. */
export function consumePdf(token: string): Buffer | null {
  const entry = store.get(token);
  if (!entry) return null;
  store.delete(token);
  if (entry.expiresAt < Date.now()) return null;
  return entry.buffer;
}
```

- [ ] **Step 2: Type-check**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "pdf-store" | head -5
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add lib/audit/pdf-store.ts
git commit -m "feat(audit): add in-memory PDF token store with TTL"
```

---

## Task 7: PDF Document

**Files:**
- Create: `lib/audit/generate-pdf.ts`

- [ ] **Step 1: Create generate-pdf.ts**

```typescript
// lib/audit/generate-pdf.ts

import React from 'react';
import {
  Document,
  Page,
  Text,
  View,
  StyleSheet,
  renderToBuffer,
} from '@react-pdf/renderer';
import type { AuditPdfData } from './types';

// ── Styles ────────────────────────────────────────────────────────────────────

const styles = StyleSheet.create({
  page: {
    fontFamily: 'Helvetica',
    fontSize: 10,
    color: '#1a1a1a',
    paddingTop: 0,
    paddingBottom: 32,
    paddingHorizontal: 0,
  },
  // Dark header banner
  header: {
    backgroundColor: '#1e3a5f',
    color: 'white',
    paddingVertical: 20,
    paddingHorizontal: 32,
    marginBottom: 24,
  },
  headerTitle: {
    fontSize: 18,
    fontFamily: 'Helvetica-Bold',
    color: 'white',
    marginBottom: 4,
  },
  headerSubtitle: {
    fontSize: 9,
    color: '#9ab3c8',
    marginBottom: 12,
  },
  headerStats: {
    flexDirection: 'row',
    gap: 24,
  },
  statBlock: {
    flexDirection: 'column',
  },
  statValue: {
    fontSize: 20,
    fontFamily: 'Helvetica-Bold',
    color: 'white',
  },
  statValueGreen: {
    fontSize: 20,
    fontFamily: 'Helvetica-Bold',
    color: '#86efac',
  },
  statValueRed: {
    fontSize: 20,
    fontFamily: 'Helvetica-Bold',
    color: '#fca5a5',
  },
  statLabel: {
    fontSize: 8,
    color: '#9ab3c8',
    marginTop: 2,
  },
  body: {
    paddingHorizontal: 32,
  },
  sectionLabel: {
    fontSize: 8,
    fontFamily: 'Helvetica-Bold',
    color: '#6b7280',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: 8,
    marginTop: 16,
  },
  groupTitle: {
    fontSize: 10,
    fontFamily: 'Helvetica-Bold',
    color: '#1e3a5f',
    borderLeftWidth: 3,
    borderLeftColor: '#1e3a5f',
    paddingLeft: 8,
    marginBottom: 4,
    marginTop: 8,
  },
  commentRow: {
    flexDirection: 'row',
    marginBottom: 4,
    paddingLeft: 12,
  },
  commentCode: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#2563eb',
    width: 28,
  },
  commentCodeFail: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#dc2626',
    width: 28,
  },
  commentText: {
    fontSize: 9,
    color: '#374151',
    flex: 1,
  },
  spagError: {
    fontSize: 8,
    color: '#dc2626',
    paddingLeft: 40,
    marginBottom: 2,
  },
  failureCard: {
    borderWidth: 1,
    borderColor: '#fecaca',
    borderRadius: 4,
    marginBottom: 6,
    overflow: 'hidden',
  },
  failureCardHeader: {
    backgroundColor: '#fee2e2',
    paddingVertical: 5,
    paddingHorizontal: 8,
    flexDirection: 'row',
    gap: 8,
  },
  failureCardHeaderText: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#991b1b',
  },
  failureCardBody: {
    paddingVertical: 5,
    paddingHorizontal: 8,
  },
  failureRule: {
    fontSize: 9,
    color: '#dc2626',
    marginBottom: 2,
  },
  untestedNote: {
    fontSize: 9,
    color: '#92400e',
    backgroundColor: '#fef3c7',
    borderRadius: 4,
    padding: 8,
    marginTop: 8,
  },
});

// ── Document component ────────────────────────────────────────────────────────

function AuditPdfDocument({ data }: { data: AuditPdfData }) {
  const passRate = data.totalReports > 0
    ? Math.round((data.passedReports / data.totalReports) * 100)
    : 0;

  const spagFailures = data.spagEntries.filter(e => !e.passed);

  // Group spag entries by group name
  const groupedEntries = new Map<string, typeof data.spagEntries>();
  for (const entry of data.spagEntries) {
    const arr = groupedEntries.get(entry.groupName) ?? [];
    arr.push(entry);
    groupedEntries.set(entry.groupName, arr);
  }

  const dateStr = data.generatedAt.toLocaleDateString('en-GB', {
    day: 'numeric', month: 'long', year: 'numeric',
  });

  return React.createElement(
    Document,
    null,
    React.createElement(
      Page,
      { size: 'A4', style: styles.page },
      // Header
      React.createElement(
        View,
        { style: styles.header },
        React.createElement(Text, { style: styles.headerTitle },
          `${data.subjectCode} — Comment Bank Audit`
        ),
        React.createElement(Text, { style: styles.headerSubtitle },
          `${data.subjectTitle} · Generated ${dateStr}`
        ),
        React.createElement(
          View,
          { style: styles.headerStats },
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, { style: styles.statValue }, String(data.totalReports)),
            React.createElement(Text, { style: styles.statLabel }, 'Reports')
          ),
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, { style: styles.statValueGreen }, `${passRate}%`),
            React.createElement(Text, { style: styles.statLabel }, 'Passed')
          ),
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, { style: styles.statValueRed }, String(spagFailures.length)),
            React.createElement(Text, { style: styles.statLabel }, 'SPAG Failures')
          ),
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, {
              style: data.untestedItems.length > 0 ? styles.statValueRed : styles.statValueGreen,
            }, String(data.untestedItems.length)),
            React.createElement(Text, { style: styles.statLabel }, 'Untested')
          )
        )
      ),
      // Body
      React.createElement(
        View,
        { style: styles.body },
        // Section 1 — Comments Audited
        React.createElement(Text, { style: styles.sectionLabel }, 'Section 1 — Comments Audited'),
        ...[...groupedEntries.entries()].map(([groupName, entries]) =>
          React.createElement(
            View,
            { key: groupName },
            React.createElement(Text, { style: styles.groupTitle }, groupName),
            ...entries.map(entry =>
              React.createElement(
                View,
                { key: entry.code },
                React.createElement(
                  View,
                  { style: styles.commentRow },
                  React.createElement(Text, {
                    style: entry.passed ? styles.commentCode : styles.commentCodeFail,
                  }, entry.passed ? entry.code : `${entry.code} ✗`),
                  React.createElement(Text, { style: styles.commentText }, entry.rawText)
                ),
                ...entry.errors.map((err, i) =>
                  React.createElement(Text, { key: i, style: styles.spagError },
                    `  ⚠ "${err.word}": ${err.message}`
                  )
                )
              )
            )
          )
        ),

        // Section 2 — Failed Standards Reports
        data.standardsFailures.length > 0
          ? React.createElement(
            View,
            null,
            React.createElement(Text, { style: styles.sectionLabel }, 'Section 2 — Failed Standards Reports'),
            ...data.standardsFailures.map(failure =>
              React.createElement(
                View,
                { key: failure.reportIndex, style: styles.failureCard },
                React.createElement(
                  View,
                  { style: styles.failureCardHeader },
                  React.createElement(Text, { style: styles.failureCardHeaderText },
                    `Report #${failure.reportIndex + 1}`
                  ),
                  React.createElement(Text, { style: styles.failureCardHeaderText },
                    Object.values(failure.codes).join(', ')
                  )
                ),
                React.createElement(
                  View,
                  { style: styles.failureCardBody },
                  ...failure.failures.map(rule =>
                    React.createElement(Text, { key: rule, style: styles.failureRule },
                      `✗ ${rule}`
                    )
                  )
                )
              )
            )
          )
          : React.createElement(Text, { style: { ...styles.sectionLabel, color: '#16a34a' } },
            '✓ All standards reports passed'
          ),

        // Untested warning
        data.untestedItems.length > 0
          ? React.createElement(Text, { style: styles.untestedNote },
            `⚠ ${data.untestedItems.length} comment code(s) were not included in any of the 50 sample reports: ` +
            data.untestedItems.map(u => `${u.code} (${u.groupName})`).join(', ')
          )
          : null
      )
    )
  );
}

// ── Public render function ────────────────────────────────────────────────────

export async function renderAuditPdf(data: AuditPdfData): Promise<Buffer> {
  const arrayBuffer = await renderToBuffer(
    React.createElement(AuditPdfDocument, { data })
  );
  return Buffer.from(arrayBuffer);
}
```

- [ ] **Step 2: Type-check**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "generate-pdf" | head -10
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add lib/audit/generate-pdf.ts
git commit -m "feat(audit): add PDF document component with dark header banner layout"
```

---

## Task 8: SSE Audit Route

**Files:**
- Create: `app/api/subjects/[subjectId]/audit/route.ts`

- [ ] **Step 1: Create the directory and route file**

```bash
mkdir -p app/api/subjects/\[subjectId\]/audit
```

```typescript
// app/api/subjects/[subjectId]/audit/route.ts

export const runtime = 'nodejs';
export const maxDuration = 300;

import { NextRequest } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';
import type { DbSubject, DbCommentGroup, DbCommentOption, DbCommonCommentGroup, DbCommonCommentOption } from '@/lib/types/db';
import type { AuditEvent, SpagAuditEntry, StandardsAuditEntry, AuditPdfData } from '@/lib/audit/types';
import { buildSampleReports } from '@/lib/audit/build-reports';
import { substituteVariables } from '@/lib/audit/substitute-variables';
import { callSpagWebhook, callStandardsWebhook } from '@/lib/audit/webhook-calls';
import { renderAuditPdf } from '@/lib/audit/generate-pdf';
import { storePdf } from '@/lib/audit/pdf-store';

export async function GET(
  _request: NextRequest,
  { params }: { params: Promise<{ subjectId: string }> }
) {
  const { subjectId } = await params;

  // ── Auth ──────────────────────────────────────────────────────────────────
  const session = await getServerSession(authOptions);
  if (!session?.user?.id) {
    return new Response('Unauthorized', { status: 401 });
  }

  // ── Fetch subject ─────────────────────────────────────────────────────────
  const { rows: subjectRows } = await pool.query<DbSubject>(
    `SELECT * FROM "Subject" WHERE id = $1`,
    [subjectId]
  );
  if (subjectRows.length === 0) {
    return new Response('Not Found', { status: 404 });
  }
  const subject = subjectRows[0];

  // ── Fetch comment groups + options (non-linked only) ──────────────────────
  const { rows: groupRows } = await pool.query<DbCommentGroup>(
    `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 AND "isLinked" = false ORDER BY "displayOrder" ASC`,
    [subjectId]
  );

  type GroupWithOptions = {
    id: string; title: string; isLinked: boolean;
    options: { id: string; code: string; text: string; displayOrder: number }[];
  };

  const subjectGroups: GroupWithOptions[] = [];
  if (groupRows.length > 0) {
    const groupIds = groupRows.map(g => g.id);
    const { rows: optRows } = await pool.query<DbCommentOption>(
      `SELECT * FROM "CommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
      [groupIds]
    );
    for (const g of groupRows) {
      subjectGroups.push({
        id: g.id,
        title: g.title,
        isLinked: g.isLinked,
        options: optRows.filter(o => o.groupId === g.id),
      });
    }
  }

  // ── Fetch common comment groups + options ─────────────────────────────────
  const { rows: ccgRows } = await pool.query<DbCommonCommentGroup>(
    `SELECT * FROM "CommonCommentGroup" WHERE "isLinked" = false ORDER BY "displayOrder" ASC`
  );
  const commonGroups: GroupWithOptions[] = [];
  if (ccgRows.length > 0) {
    const ccgIds = ccgRows.map(g => g.id);
    const { rows: ccoRows } = await pool.query<DbCommonCommentOption>(
      `SELECT * FROM "CommonCommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
      [ccgIds]
    );
    for (const g of ccgRows) {
      commonGroups.push({
        id: g.id,
        title: g.title,
        isLinked: g.isLinked,
        options: ccoRows.filter(o => o.groupId === g.id),
      });
    }
  }

  // ── Fetch ignored words for this teacher ──────────────────────────────────
  const { rows: ignoredRows } = await pool.query<{ word: string }>(
    `SELECT word FROM "IgnoredWord" WHERE "teacherId" = $1`,
    [session.user.id]
  );
  const ignoredWords = new Set(ignoredRows.map(r => r.word.toLowerCase()));

  // ── Collect all comment options for SPAG Phase 1 ─────────────────────────
  const allComments: { code: string; text: string; groupName: string }[] = [];
  for (const g of subjectGroups) {
    for (const opt of g.options) {
      allComments.push({ code: opt.code, text: opt.text, groupName: g.title });
    }
  }
  for (const g of commonGroups) {
    for (const opt of g.options) {
      allComments.push({ code: opt.code, text: opt.text, groupName: g.title });
    }
  }

  // ── Build 50 sample reports ───────────────────────────────────────────────
  const { reports, untestedItems } = buildSampleReports(
    subjectGroups, commonGroups, subject.title ?? subject.code
  );

  // ── Stream ────────────────────────────────────────────────────────────────
  const encoder = new TextEncoder();

  const stream = new ReadableStream({
    async start(controller) {
      // EventSource requires SSE format: "data: <json>\n\n"
      const send = (event: AuditEvent) => {
        controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
      };

      try {
        // Init
        send({ type: 'init', totalComments: allComments.length, totalReports: reports.length });

        // ── Phase 1: SPAG ──────────────────────────────────────────────────
        const spagEntries: SpagAuditEntry[] = [];

        for (const comment of allComments) {
          // Substitute variables before sending to SPAG — raw <Name> etc. would be flagged as errors
          const substituted = substituteVariables(comment.text, subject.title ?? subject.code);
          const { passed, errors } = await callSpagWebhook(substituted, ignoredWords);

          const entry: SpagAuditEntry = {
            code: comment.code,
            groupName: comment.groupName,
            rawText: comment.text,
            passed,
            errors,
          };
          spagEntries.push(entry);

          send({
            type: 'spag',
            code: comment.code,
            groupName: comment.groupName,
            passed,
            errors,
          });
        }

        send({ type: 'spag_done' });

        // ── Phase 2: Standards ─────────────────────────────────────────────
        const standardsFailures: StandardsAuditEntry[] = [];
        let passedReports = 0;

        for (const report of reports) {
          const { passed, failures } = await callStandardsWebhook(report.assembledText);
          if (passed) passedReports++;

          const codes: Record<string, string> = {};
          for (const [groupId, sel] of Object.entries(report.selections)) {
            codes[groupId] = sel.code;
          }

          if (!passed) {
            standardsFailures.push({ reportIndex: report.reportIndex, codes, passed, failures });
          }

          send({
            type: 'standards',
            reportIndex: report.reportIndex,
            codes,
            passed,
            failures,
          });
        }

        send({ type: 'standards_done' });
        send({ type: 'untested', items: untestedItems });

        // ── Generate PDF ───────────────────────────────────────────────────
        const groupTitles: Record<string, string> = {};
        for (const g of subjectGroups) groupTitles[g.id] = g.title;

        const pdfData: AuditPdfData = {
          subjectTitle: subject.title ?? subject.code,
          subjectCode: subject.code,
          generatedAt: new Date(),
          totalReports: reports.length,
          passedReports,
          spagEntries,
          standardsFailures,
          untestedItems,
          groupTitles,
        };

        const pdfBuffer = await renderAuditPdf(pdfData);
        const token = storePdf(pdfBuffer);

        send({ type: 'complete', pdfUrl: `/api/subjects/${subjectId}/audit/pdf?token=${token}` });
      } catch (err) {
        send({ type: 'error', message: err instanceof Error ? err.message : String(err) });
      } finally {
        controller.close();
      }
    },
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache, no-transform',
      'Connection': 'keep-alive',
      'X-Accel-Buffering': 'no',
    },
  });
}
```

- [ ] **Step 2: Type-check**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "audit/route" | head -10
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add app/api/subjects/\[subjectId\]/audit/route.ts
git commit -m "feat(audit): add SSE streaming audit route"
```

---

## Task 9: PDF Download Route

**Files:**
- Create: `app/api/subjects/[subjectId]/audit/pdf/route.ts`

- [ ] **Step 1: Create the route**

```bash
mkdir -p "app/api/subjects/[subjectId]/audit/pdf"
```

```typescript
// app/api/subjects/[subjectId]/audit/pdf/route.ts

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

  return new Response(buffer, {
    headers: {
      'Content-Type': 'application/pdf',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': String(buffer.length),
    },
  });
}
```

- [ ] **Step 2: Type-check**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "audit/pdf" | head -10
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add "app/api/subjects/[subjectId]/audit/pdf/route.ts"
git commit -m "feat(audit): add token-gated PDF download route"
```

---

## Task 10: AuditModal Component

**Files:**
- Create: `components/AuditModal.tsx`

- [ ] **Step 1: Create AuditModal.tsx**

```typescript
// components/AuditModal.tsx
'use client';

import { useEffect, useRef, useState, useCallback } from 'react';
import { createPortal } from 'react-dom';
import type { AuditEvent } from '@/lib/audit/types';

type Phase =
  | { name: 'idle' }
  | {
      name: 'phase1';
      totalComments: number;
      checkedComments: number;
      currentLabel: string;
    }
  | {
      name: 'phase2';
      totalComments: number;
      totalReports: number;
      checkedReports: number;
      currentLabel: string;
    }
  | {
      name: 'complete';
      totalReports: number;
      passedReports: number;
      spagFailures: number;
      untestedCount: number;
      pdfUrl: string;
    }
  | { name: 'error'; message: string };

interface AuditModalProps {
  subjectId: string;
  subjectTitle: string;
  isOpen: boolean;
  onClose: () => void;
}

export default function AuditModal({
  subjectId,
  subjectTitle,
  isOpen,
  onClose,
}: AuditModalProps) {
  const [phase, setPhase] = useState<Phase>({ name: 'idle' });
  const esRef = useRef<EventSource | null>(null);
  // Track counts for complete summary
  const spagFailRef = useRef(0);
  const totalCommentsRef = useRef(0);
  const totalReportsRef = useRef(0);
  const passedReportsRef = useRef(0);

  const closeStream = useCallback(() => {
    esRef.current?.close();
    esRef.current = null;
  }, []);

  useEffect(() => {
    if (!isOpen) {
      closeStream();
      setPhase({ name: 'idle' });
      spagFailRef.current = 0;
      totalCommentsRef.current = 0;
      totalReportsRef.current = 0;
      passedReportsRef.current = 0;
      return;
    }

    // Start stream
    const es = new EventSource(`/api/subjects/${subjectId}/audit`);
    esRef.current = es;

    es.onmessage = (e: MessageEvent) => {
      const event = JSON.parse(e.data) as AuditEvent;
      handleEvent(event);
    };

    es.onerror = () => {
      closeStream();
      setPhase({ name: 'error', message: 'Connection to audit service lost.' });
    };

    return () => {
      closeStream();
    };
  }, [isOpen, subjectId, closeStream]);

  // Separate event handler to allow updating untestedCount on complete
  const untestedCountRef = useRef(0);

  function handleEvent(event: AuditEvent) {
    switch (event.type) {
      case 'init':
        totalCommentsRef.current = event.totalComments;
        totalReportsRef.current = event.totalReports;
        setPhase({
          name: 'phase1',
          totalComments: event.totalComments,
          checkedComments: 0,
          currentLabel: 'Starting…',
        });
        break;

      case 'spag':
        if (!event.passed) spagFailRef.current++;
        setPhase(prev =>
          prev.name === 'phase1'
            ? { ...prev, checkedComments: prev.checkedComments + 1, currentLabel: `${event.groupName}: ${event.code}` }
            : prev
        );
        break;

      case 'spag_done':
        setPhase(prev =>
          prev.name === 'phase1'
            ? { name: 'phase2', totalComments: prev.totalComments, totalReports: totalReportsRef.current, checkedReports: 0, currentLabel: 'Building sample reports…' }
            : prev
        );
        break;

      case 'standards':
        if (event.passed) passedReportsRef.current++;
        setPhase(prev =>
          prev.name === 'phase2'
            ? { ...prev, checkedReports: prev.checkedReports + 1, currentLabel: `Report #${event.reportIndex + 1}` }
            : prev
        );
        break;

      case 'untested':
        untestedCountRef.current = event.items.length;
        break;

      case 'complete':
        closeStream();
        setPhase({
          name: 'complete',
          totalReports: totalReportsRef.current,
          passedReports: passedReportsRef.current,
          spagFailures: spagFailRef.current,
          untestedCount: untestedCountRef.current,
          pdfUrl: event.pdfUrl,
        });
        break;

      case 'error':
        closeStream();
        setPhase({ name: 'error', message: event.message });
        break;
    }
  }

  // Escape key
  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isOpen) {
        closeStream();
        onClose();
      }
    };
    document.addEventListener('keydown', handleKey);
    return () => document.removeEventListener('keydown', handleKey);
  }, [isOpen, onClose, closeStream]);

  if (!isOpen || typeof document === 'undefined') return null;

  const handleDownload = () => {
    if (phase.name !== 'complete') return;
    const a = document.createElement('a');
    a.href = phase.pdfUrl;
    a.download = '';
    a.click();
  };

  const handleCancel = () => {
    closeStream();
    onClose();
  };

  const handleRetry = () => {
    setPhase({ name: 'idle' });
    spagFailRef.current = 0;
    totalCommentsRef.current = 0;
    totalReportsRef.current = 0;
    passedReportsRef.current = 0;
    untestedCountRef.current = 0;
    // Re-trigger by closing and re-opening — parent controls isOpen
    onClose();
  };

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-center justify-center">
      <div className="absolute inset-0 bg-black/50 backdrop-blur-sm" onClick={handleCancel} />
      <div className="relative bg-white dark:bg-gray-900 rounded-xl shadow-2xl max-w-md w-full mx-4 p-6">
        <h3 className="text-lg font-bold text-gray-900 dark:text-white mb-1">
          Comment Bank Audit
        </h3>
        <p className="text-sm text-gray-500 dark:text-gray-400 mb-6">{subjectTitle}</p>

        {phase.name === 'idle' && (
          <p className="text-sm text-gray-500">Starting audit…</p>
        )}

        {(phase.name === 'phase1' || phase.name === 'phase2') && (
          <div className="space-y-5">
            {/* Phase 1 bar */}
            <div>
              <div className="flex justify-between text-xs text-gray-600 dark:text-gray-400 mb-1">
                <span>Phase 1: SPAG checking comments</span>
                <span>
                  {phase.name === 'phase1' ? phase.checkedComments : (phase as { totalComments: number }).totalComments}
                  {' / '}
                  {phase.name === 'phase1' ? phase.totalComments : (phase as { totalComments: number }).totalComments}
                </span>
              </div>
              <div className="bg-gray-200 dark:bg-gray-700 rounded-full h-2">
                <div
                  className="bg-blue-500 h-2 rounded-full transition-all duration-300"
                  style={{
                    width: phase.name === 'phase1' && phase.totalComments > 0
                      ? `${(phase.checkedComments / phase.totalComments) * 100}%`
                      : '100%',
                  }}
                />
              </div>
            </div>

            {/* Phase 2 bar */}
            <div className={phase.name === 'phase1' ? 'opacity-40' : ''}>
              <div className="flex justify-between text-xs text-gray-600 dark:text-gray-400 mb-1">
                <span>Phase 2: Standards checking reports</span>
                <span>
                  {phase.name === 'phase2' ? phase.checkedReports : 0}
                  {' / '}
                  {phase.name === 'phase2' ? phase.totalReports : totalReportsRef.current}
                </span>
              </div>
              <div className="bg-gray-200 dark:bg-gray-700 rounded-full h-2">
                <div
                  className="bg-blue-500 h-2 rounded-full transition-all duration-300"
                  style={{
                    width: phase.name === 'phase2' && phase.totalReports > 0
                      ? `${(phase.checkedReports / phase.totalReports) * 100}%`
                      : '0%',
                  }}
                />
              </div>
            </div>

            <p className="text-xs text-gray-400 dark:text-gray-500 truncate">
              {phase.name === 'phase1' ? phase.currentLabel : (phase as { currentLabel: string }).currentLabel}
            </p>

            <div className="flex justify-end">
              <button
                onClick={handleCancel}
                className="px-4 py-2 text-sm text-gray-600 dark:text-gray-400 hover:text-gray-800 dark:hover:text-white transition-colors"
              >
                Cancel
              </button>
            </div>
          </div>
        )}

        {phase.name === 'complete' && (
          <div className="space-y-4">
            <div className="flex items-center gap-2 text-green-600 dark:text-green-400">
              <span className="material-symbols-outlined">check_circle</span>
              <span className="font-semibold">Audit Complete</span>
            </div>
            <ul className="text-sm text-gray-600 dark:text-gray-400 space-y-1">
              <li>{phase.totalReports} reports generated</li>
              <li>
                {phase.totalReports > 0
                  ? Math.round((phase.passedReports / phase.totalReports) * 100)
                  : 0}
                % passed standards checks
              </li>
              <li className={phase.spagFailures > 0 ? 'text-red-500' : ''}>
                {phase.spagFailures} SPAG {phase.spagFailures === 1 ? 'failure' : 'failures'}
              </li>
              <li className={phase.untestedCount > 0 ? 'text-amber-500' : ''}>
                {phase.untestedCount} comment {phase.untestedCount === 1 ? 'code' : 'codes'} untested
              </li>
            </ul>
            <div className="flex gap-3 pt-2">
              <button
                onClick={onClose}
                className="flex-1 px-4 py-2 text-sm border border-gray-300 dark:border-gray-600 rounded-lg text-gray-700 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-800 transition-colors"
              >
                Close
              </button>
              <button
                onClick={handleDownload}
                className="flex-1 px-4 py-2 text-sm bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors flex items-center justify-center gap-2"
              >
                <span className="material-symbols-outlined text-base">download</span>
                Download PDF
              </button>
            </div>
          </div>
        )}

        {phase.name === 'error' && (
          <div className="space-y-4">
            <div className="flex items-center gap-2 text-red-500">
              <span className="material-symbols-outlined">error</span>
              <span className="font-semibold">Audit Failed</span>
            </div>
            <p className="text-sm text-gray-600 dark:text-gray-400">{phase.message}</p>
            <div className="flex gap-3">
              <button
                onClick={onClose}
                className="flex-1 px-4 py-2 text-sm border border-gray-300 dark:border-gray-600 rounded-lg text-gray-700 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-800 transition-colors"
              >
                Close
              </button>
              <button
                onClick={handleRetry}
                className="flex-1 px-4 py-2 text-sm bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                Retry
              </button>
            </div>
          </div>
        )}
      </div>
    </div>,
    document.body
  );
}
```

- [ ] **Step 2: Type-check**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "AuditModal" | head -10
```

Expected: no errors.

- [ ] **Step 3: Commit**

```bash
git add components/AuditModal.tsx
git commit -m "feat(audit): add AuditModal component with SSE client and dual progress bars"
```

---

## Task 11: Wire Up Audit Button on Subject Page

**Files:**
- Modify: `app/hod/subject/[subjectId]/page.tsx`

The subject page is a server component. Add a small `'use client'` wrapper that holds the `isOpen` state and renders the `AuditModal`.

- [ ] **Step 1: Create AuditButton wrapper component**

Create `app/hod/subject/[subjectId]/_components/audit-button.tsx`:

```typescript
// app/hod/subject/[subjectId]/_components/audit-button.tsx
'use client';

import { useState } from 'react';
import AuditModal from '@/components/AuditModal';

interface AuditButtonProps {
  subjectId: string;
  subjectTitle: string;
}

export function AuditButton({ subjectId, subjectTitle }: AuditButtonProps) {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <>
      <button
        onClick={() => setIsOpen(true)}
        className="flex items-center gap-2 px-4 py-2 bg-purple-100 dark:bg-purple-900/30 text-purple-700 dark:text-purple-400 rounded-lg hover:bg-purple-200 dark:hover:bg-purple-900/50 transition-colors font-medium text-sm"
      >
        <span className="material-symbols-outlined text-lg">fact_check</span>
        Audit Comments
      </button>

      <AuditModal
        subjectId={subjectId}
        subjectTitle={subjectTitle}
        isOpen={isOpen}
        onClose={() => setIsOpen(false)}
      />
    </>
  );
}
```

- [ ] **Step 2: Import and add AuditButton to the subject page**

In `app/hod/subject/[subjectId]/page.tsx`, add the import after existing imports:

```typescript
import { AuditButton } from './_components/audit-button';
```

Then inside the `<div className="flex items-center gap-4">` block (around line 93), add the button before the review link:

```tsx
<AuditButton
  subjectId={subjectId}
  subjectTitle={`${subject.code} — ${subject.title}`}
/>
<div className="h-10 w-px bg-gray-200 dark:bg-gray-700 hidden md:block"></div>
```

- [ ] **Step 3: Full type-check**

```bash
npx tsc --noEmit --incremental false 2>&1 | head -20
```

Expected: no errors.

- [ ] **Step 4: Run all tests**

```bash
npx vitest run lib/__tests__/audit-substitute-variables.test.ts lib/__tests__/audit-build-reports.test.ts
```

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add app/hod/subject/\[subjectId\]/_components/audit-button.tsx app/hod/subject/\[subjectId\]/page.tsx
git commit -m "feat(audit): add Audit Comments button to subject admin page"
```

---

## Manual Smoke Test

After all tasks are complete:

1. Start the dev server: `npm run dev`
2. Navigate to any subject admin page: `http://localhost:3000/hod/subject/<id>`
3. Click **Audit Comments** — the modal should appear
4. Both progress bars should be visible; Phase 2 bar should be greyed out initially
5. As SPAG events arrive, Phase 1 bar should fill and the current label should update
6. When Phase 1 completes, Phase 2 bar should activate and fill
7. On completion, the summary view shows stats and a **Download PDF** button
8. Clicking **Download PDF** should download a PDF with a dark navy header, subject stats, and grouped comment sections
