'use client';

import { useState, useMemo } from 'react';
import type { AiSuggestion, DiffSegment } from '@/lib/types/ai-check';
import { groupChanges } from '@/lib/utils/diff';

interface AiSuggestionPanelProps {
  suggestion: AiSuggestion;
  onApplyChange: (newText: string, isLastChange: boolean) => void;
  onApplyAll: (improvedText: string) => void;
  onDismiss: () => void;
}

function humanizeKey(key: string): string {
  return key
    .replace(/_/g, ' ')
    .replace(/([A-Z])/g, ' $1')
    .trim();
}

function buildText(segments: DiffSegment[], appliedIds: Set<number>): string {
  return segments
    .map(seg => {
      if (seg.type === 'unchanged') return seg.tokens.map(t => t.text).join('');
      if (appliedIds.has(seg.group.id)) return seg.group.added.map(t => t.text).join('');
      return seg.group.removed.map(t => t.text).join('');
    })
    .join('');
}

export default function AiSuggestionPanel({ suggestion, onApplyChange, onApplyAll, onDismiss }: AiSuggestionPanelProps) {
  const segments = useMemo(() => groupChanges(suggestion.diff), [suggestion.diff]);
  const [appliedIds, setAppliedIds] = useState<Set<number>>(new Set());

  const changeGroups = useMemo(() => segments.filter(s => s.type === 'change'), [segments]);
  const totalChanges = changeGroups.length;
  const remainingGroups = changeGroups.filter(s => s.type === 'change' && !appliedIds.has((s as Extract<DiffSegment, {type:'change'}>).group.id));

  const ruleEntries = Object.entries(suggestion.ruleChecks);
  const passedCount = ruleEntries.filter(([, v]) => v).length;

  const handleApplyChange = (id: number) => {
    const newApplied = new Set([...appliedIds, id]);
    setAppliedIds(newApplied);
    const isLast = changeGroups.every(
      s => s.type === 'change' && newApplied.has(s.group.id)
    );
    onApplyChange(buildText(segments, newApplied), isLast);
  };

  const handleApplyAll = () => {
    const allIds = new Set(
      changeGroups
        .filter(s => s.type === 'change')
        .map(s => (s as Extract<DiffSegment, {type:'change'}>).group.id)
    );
    onApplyAll(buildText(segments, allIds));
  };

  return (
    <div className="mx-8 mb-4 bg-white dark:bg-gray-900 border border-blue-200 dark:border-blue-800 rounded-lg overflow-hidden">

      {/* Header */}
      <div className="bg-blue-50 dark:bg-blue-900/20 border-b border-blue-200 dark:border-blue-800 px-4 py-3 flex items-center gap-2">
        <span className="material-symbols-outlined text-blue-600 dark:text-blue-400 text-base">auto_fix_high</span>
        <strong className="text-blue-700 dark:text-blue-400 text-sm">AI Suggestion</strong>
        <span className="ml-auto text-xs text-gray-500 dark:text-gray-400 italic">
          {remainingGroups.length > 0
            ? `${remainingGroups.length} of ${totalChanges} changes remaining — click to apply`
            : 'All changes applied'}
        </span>
      </div>

      {/* Diff section */}
      <div className="px-4 py-3 border-b border-gray-100 dark:border-gray-800">
        <div className="bg-gray-50 dark:bg-gray-950 border border-gray-200 dark:border-gray-700 rounded-lg p-3 text-base leading-loose font-display">
          {segments.map((seg, i) => {
            if (seg.type === 'unchanged') {
              return <span key={i}>{seg.tokens.map(t => t.text).join('')}</span>;
            }
            const { group } = seg;

            if (appliedIds.has(group.id)) {
              // Applied — show new text, no interaction
              return (
                <span key={i} className="text-green-700 dark:text-green-400">
                  {group.added.map(t => t.text).join('')}
                </span>
              );
            }

            // Unapplied — show old→new, click to apply
            return (
              <span
                key={i}
                onClick={() => handleApplyChange(group.id)}
                className="inline cursor-pointer"
                title="Click to apply this change"
              >
                {group.removed.length > 0 && (
                  <span className="bg-red-50 dark:bg-red-900/30 text-red-600 dark:text-red-400 line-through rounded px-1 border border-red-200 dark:border-red-800 mx-0.5">
                    {group.removed.map(t => t.text).join('')}
                  </span>
                )}
                {group.removed.length > 0 && group.added.length > 0 && (
                  <span className="text-gray-400 text-xs mx-0.5">→</span>
                )}
                {group.added.length > 0 && (
                  <span className="bg-green-50 dark:bg-green-900/30 text-green-700 dark:text-green-400 rounded px-1 border border-dashed border-green-300 dark:border-green-700 mx-0.5">
                    {group.added.map(t => t.text).join('')}
                  </span>
                )}
              </span>
            );
          })}
        </div>
        {remainingGroups.length > 0 && (
          <p className="mt-2 text-xs text-gray-400 flex items-center gap-1">
            <span className="material-symbols-outlined text-sm">touch_app</span>
            Click any highlighted change to apply it individually
          </p>
        )}
      </div>

      {/* Rule checks section */}
      {ruleEntries.length > 0 && (
        <div className="px-4 py-3 border-b border-gray-100 dark:border-gray-800">
          <p className="text-[10px] font-bold uppercase tracking-widest text-gray-400 mb-2">
            Rule checks — {passedCount} / {ruleEntries.length} passed
          </p>
          <div className="grid grid-cols-2 gap-1">
            {ruleEntries.map(([key, passed]) => (
              <div
                key={key}
                className={`flex items-center gap-2 px-2 py-1.5 rounded text-xs ${
                  passed
                    ? 'bg-green-50 dark:bg-green-900/20 text-green-700 dark:text-green-400'
                    : 'bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-400 border-l-2 border-red-400'
                }`}
              >
                <span className="font-bold flex-shrink-0">{passed ? '✓' : '✗'}</span>
                <span>{humanizeKey(key)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Actions */}
      <div className="px-4 py-3 flex items-center gap-2">
        {remainingGroups.length > 0 && (
          <button
            onClick={handleApplyAll}
            className="flex items-center gap-1.5 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg transition-colors"
          >
            <span className="material-symbols-outlined text-sm">done_all</span>
            Apply all
          </button>
        )}
        <button
          onClick={onDismiss}
          className="flex items-center gap-1.5 px-3 py-2 ml-auto bg-white dark:bg-gray-900 hover:bg-gray-50 dark:hover:bg-gray-800 text-gray-600 dark:text-gray-400 text-xs font-medium rounded-lg border border-gray-200 dark:border-gray-700 transition-colors"
        >
          <span className="material-symbols-outlined text-sm">close</span>
          Dismiss
        </button>
      </div>
    </div>
  );
}
