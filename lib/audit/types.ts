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
  pdfBase64: string;
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
};
