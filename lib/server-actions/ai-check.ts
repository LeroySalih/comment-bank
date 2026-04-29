'use server';

import type { AiSuggestion } from '@/lib/types/ai-check';
import { computeDiff } from '@/lib/utils/diff';

// Substitution map for the mock — simulates vocabulary improvements
const MOCK_SUBSTITUTIONS: Record<string, string> = {
  good: 'excellent',
  well: 'effectively',
  hard: 'diligently',
  bad: 'challenging',
  big: 'significant',
  shows: 'demonstrates',
  show: 'demonstrate',
  works: 'strives',
  work: 'strive',
  tries: 'endeavours',
  try: 'endeavour',
  nice: 'commendable',
  improve: 'enhance',
  improving: 'enhancing',
  improved: 'enhanced',
  help: 'support',
  helped: 'supported',
  helps: 'supports',
  make: 'achieve',
  makes: 'achieves',
  made: 'achieved',
};

function mockImprove(text: string): string {
  return text
    .split(/\b/)
    .map(word => {
      const lower = word.toLowerCase();
      const sub = MOCK_SUBSTITUTIONS[lower];
      if (!sub) return word;
      // Preserve original capitalisation
      if (word[0] === word[0].toUpperCase()) {
        return sub.charAt(0).toUpperCase() + sub.slice(1);
      }
      return sub;
    })
    .join('');
}

export async function requestAiCheck(
  assignmentId: string,
  commentText: string
): Promise<{ success: true; suggestion: AiSuggestion } | { success: false; error: string }> {
  try {
    if (!commentText.trim()) {
      return { success: false, error: 'No comment text to check' };
    }

    // TODO: replace mockImprove with real AI call
    const improved = mockImprove(commentText);
    const diff = computeDiff(commentText, improved);

    return {
      success: true,
      suggestion: {
        original: commentText,
        improved,
        diff,
      },
    };
  } catch (err) {
    return { success: false, error: 'AI check failed' };
  }
}
