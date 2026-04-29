export type DiffTokenType = 'unchanged' | 'added' | 'removed';

export type DiffToken = {
  text: string;
  type: DiffTokenType;
};

export type AiSuggestion = {
  original: string;
  improved: string;
  diff: DiffToken[];
};
