'use server';

import type { AiSuggestion } from '@/lib/types/ai-check';
import { computeDiff } from '@/lib/utils/diff';

const AI_WEBHOOK_URL = 'https://n8n.mr-salih.org/webhook-test/comment-bank/ai-suggestion';

/** Remove [cite: N] and [cite: N, M, ...] references injected by the AI workflow */
function stripCitations(text: string): string {
  return text.replace(/\s*\[cite:[^\]]+\]/g, '');
}

export async function requestAiCheck(
  assignmentId: string,
  commentText: string
): Promise<{ success: true; suggestion: AiSuggestion } | { success: false; error: string }> {
  if (!commentText.trim()) {
    return { success: false, error: 'No comment text to check' };
  }

  let improved: string;
  let ruleChecks: Record<string, boolean> = {};

  try {
    const response = await fetch(AI_WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ comment: commentText }),
    });

    if (!response.ok) {
      return { success: false, error: `AI service returned ${response.status}` };
    }

    const data = await response.json();

    if (typeof data?.improved !== 'string') {
      return { success: false, error: 'Unexpected response from AI service' };
    }

    improved = stripCitations(data.improved);

    if (data.rule_checks && typeof data.rule_checks === 'object') {
      ruleChecks = data.rule_checks as Record<string, boolean>;
    }
  } catch {
    return { success: false, error: 'Could not reach AI service' };
  }

  const diff = computeDiff(commentText, improved);

  return {
    success: true,
    suggestion: {
      original: commentText,
      improved,
      diff,
      ruleChecks,
    },
  };
}
