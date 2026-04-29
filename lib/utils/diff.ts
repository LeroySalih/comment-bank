// lib/utils/diff.ts
import type { DiffToken, DiffSegment } from '@/lib/types/ai-check';

function tokenize(text: string): string[] {
  // Split on whitespace, preserving punctuation attached to words
  return text.match(/\S+|\s+/g) ?? [];
}

function lcs(a: string[], b: string[]): number[][] {
  const m = a.length;
  const n = b.length;
  const dp: number[][] = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a[i - 1] === b[j - 1] ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1]);
    }
  }
  return dp;
}

export function computeDiff(original: string, improved: string): DiffToken[] {
  const a = tokenize(original);
  const b = tokenize(improved);
  const dp = lcs(a, b);
  const tokens: DiffToken[] = [];

  let i = a.length;
  let j = b.length;
  const path: DiffToken[] = [];

  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && a[i - 1] === b[j - 1]) {
      path.push({ text: a[i - 1], type: 'unchanged' });
      i--;
      j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
      path.push({ text: b[j - 1], type: 'added' });
      j--;
    } else {
      path.push({ text: a[i - 1], type: 'removed' });
      i--;
    }
  }

  return path.reverse();
}

export function groupChanges(diff: DiffToken[]): DiffSegment[] {
  const segments: DiffSegment[] = [];
  let changeId = 0;
  let i = 0;

  while (i < diff.length) {
    if (diff[i].type === 'unchanged') {
      const tokens: DiffToken[] = [];
      while (i < diff.length && diff[i].type === 'unchanged') {
        tokens.push(diff[i]);
        i++;
      }
      segments.push({ type: 'unchanged', tokens });
    } else {
      const removed: DiffToken[] = [];
      const added: DiffToken[] = [];
      while (i < diff.length && diff[i].type !== 'unchanged') {
        if (diff[i].type === 'removed') removed.push(diff[i]);
        else added.push(diff[i]);
        i++;
      }
      segments.push({ type: 'change', group: { id: changeId++, removed, added } });
    }
  }

  return segments;
}
