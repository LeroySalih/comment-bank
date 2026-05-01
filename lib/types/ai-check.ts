export type DiffTokenType = 'unchanged' | 'added' | 'removed';

export type DiffToken = {
  text: string;
  type: DiffTokenType;
};

export type DiffChangeGroup = {
  id: number;
  removed: DiffToken[];
  added: DiffToken[];
};

export type DiffSegment =
  | { type: 'unchanged'; tokens: DiffToken[] }
  | { type: 'change'; group: DiffChangeGroup };

export type AiSuggestion = {
  original: string;
  improved: string;
  diff: DiffToken[];
  ruleChecks: Record<string, boolean>;
};

export type SpagMatch = {
  word: string;
  offset: number;
  length: number;
  replacements: string[];
  message: string;
};

export type ResolvedSpagMatch =
  | { match: SpagMatch; action: 'accepted'; replacement: string }
  | { match: SpagMatch; action: 'ignored' };

export type SpagResult = {
  matches: SpagMatch[];
};

// Standards (second-step) check — pass/fail booleans per rule
export type StandardsRuleKey =
  | 'UKSpelling'
  | 'CourseOverviewIncluded'
  | 'AcademicPerformanceIncluded'
  | 'TargetWordCountMet'
  | 'TerminologyCorrect'
  | 'JargonFree'
  | 'DataSpecific'
  | 'ToneBalanced'
  | 'SocialSkillsIncluded'
  | 'CollaborationIncluded'
  | 'BehaviourIncluded'
  | 'ParentalSupportIncluded'
  | 'Formatting';

export type StandardsRuleEntry = {
  result: boolean;
  instances?: string[];
  wordCount?: number;
};

export type StandardsResult = {
  Status: StandardsRuleEntry;
} & Record<StandardsRuleKey, StandardsRuleEntry>;
