'use server';

import type { AiSuggestion } from '@/lib/types/ai-check';
import { computeDiff } from '@/lib/utils/diff';

const AI_WEBHOOK_URL = 'https://n8n.mr-salih.org/webhook-test/comment-bank/ai-suggestion';

export async function requestAiCheck(
  assignmentId: string,
  commentText: string
): Promise<{ success: true; suggestion: AiSuggestion } | { success: false; error: string }> {
  if (!commentText.trim()) {
    return { success: false, error: 'No comment text to check' };
  }

  let improved: string;

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

    improved = data.improved;
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
    },
  };
}
