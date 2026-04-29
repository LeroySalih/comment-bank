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
