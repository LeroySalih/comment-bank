'use client';

import { useState, useMemo } from 'react';
import type { AiSuggestion, DiffSegment } from '@/lib/types/ai-check';
import { groupChanges } from '@/lib/utils/diff';

interface AiSuggestionPanelProps {
  suggestion: AiSuggestion;
  onAccept: (improved: string) => void;
  onDismiss: () => void;
}

type ChangeState = 'pending' | 'accepted';

function humanizeKey(key: string): string {
  return key
    .replace(/_/g, ' ')
    .replace(/([A-Z])/g, ' $1')
    .trim();
}

function buildFinalText(segments: DiffSegment[], states: Record<number, ChangeState>): string {
  return segments
    .map(seg => {
      if (seg.type === 'unchanged') return seg.tokens.map(t => t.text).join('');
      const state = states[seg.group.id] ?? 'pending';
      if (state === 'accepted') return seg.group.added.map(t => t.text).join('');
      return seg.group.removed.map(t => t.text).join('');
    })
    .join('');
}

export default function AiSuggestionPanel({ suggestion, onAccept, onDismiss }: AiSuggestionPanelProps) {
  const segments = useMemo(() => groupChanges(suggestion.diff), [suggestion.diff]);
  const [changeStates, setChangeStates] = useState<Record<number, ChangeState>>({});

  const changeGroups = segments.filter(s => s.type === 'change');
  const totalChanges = changeGroups.length;
  const acceptedCount = Object.values(changeStates).filter(s => s === 'accepted').length;

  const ruleEntries = Object.entries(suggestion.ruleChecks);
  const passedCount = ruleEntries.filter(([, v]) => v).length;

  const handleToggle = (id: number) => {
    setChangeStates(prev => ({
      ...prev,
      [id]: prev[id] === 'accepted' ? 'pending' : 'accepted',
    }));
  };

  const handleAcceptAll = () => {
    const all: Record<number, ChangeState> = {};
    segments.forEach(seg => {
      if (seg.type === 'change') all[seg.group.id] = 'accepted';
    });
    setChangeStates(all);
  };

  const handleApply = () => {
    onAccept(buildFinalText(segments, changeStates));
  };

  return (
    <div className="mx-8 mb-4 bg-white dark:bg-gray-900 border border-blue-200 dark:border-blue-800 rounded-lg overflow-hidden">

      {/* Header */}
      <div className="bg-blue-50 dark:bg-blue-900/20 border-b border-blue-200 dark:border-blue-800 px-4 py-3 flex items-center gap-2">
        <span className="material-symbols-outlined text-blue-600 dark:text-blue-400 text-base">auto_fix_high</span>
        <strong className="text-blue-700 dark:text-blue-400 text-sm">AI Suggestion</strong>
        <span className="ml-auto text-xs text-gray-500 dark:text-gray-400 italic">Click a change to accept it</span>
      </div>

      {/* Diff section */}
      <div className="px-4 py-3 border-b border-gray-100 dark:border-gray-800">
        <p className="text-[10px] font-bold uppercase tracking-widest text-gray-400 mb-2">
          Suggested changes — {acceptedCount} of {totalChanges} accepted
        </p>
        <div className="bg-gray-50 dark:bg-gray-950 border border-gray-200 dark:border-gray-700 rounded-lg p-3 text-base leading-loose font-display">
          {segments.map((seg, i) => {
            if (seg.type === 'unchanged') {
              return <span key={i}>{seg.tokens.map(t => t.text).join('')}</span>;
            }
            const { group } = seg;
            const state = changeStates[group.id] ?? 'pending';
            if (state === 'accepted') {
              return (
                <span
                  key={i}
                  onClick={() => handleToggle(group.id)}
                  className="bg-green-100 dark:bg-green-900/40 text-green-700 dark:text-green-400 border border-green-300 dark:border-green-700 rounded px-1 cursor-pointer"
                  title="Accepted — click to undo"
                >
                  {group.added.map(t => t.text).join('')} ✓
                </span>
              );
            }
            return (
              <span
                key={i}
                onClick={() => handleToggle(group.id)}
                className="bg-amber-50 dark:bg-amber-900/20 border border-dashed border-amber-400 dark:border-amber-600 rounded px-1 cursor-pointer"
                title="Click to accept this change"
              >
                <span className="text-red-600 dark:text-red-400 line-through">{group.removed.map(t => t.text).join('')}</span>
                {group.removed.length > 0 && group.added.length > 0 && ' '}
                <span className="text-green-700 dark:text-green-400">{group.added.map(t => t.text).join('')}</span>
              </span>
            );
          })}
        </div>
        <div className="flex items-center gap-4 mt-2 text-xs text-gray-400">
          <span className="flex items-center gap-1">
            <span className="w-3 h-3 rounded border border-dashed border-amber-400 bg-amber-50 inline-block"></span>
            Pending — click to accept
          </span>
          <span className="flex items-center gap-1">
            <span className="w-3 h-3 rounded border border-green-300 bg-green-100 inline-block"></span>
            Accepted
          </span>
        </div>
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
        <button
          onClick={handleAcceptAll}
          className="flex items-center gap-1.5 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold rounded-lg transition-colors"
        >
          <span className="material-symbols-outlined text-sm">done_all</span>
          Accept all changes
        </button>
        <button
          onClick={handleApply}
          disabled={acceptedCount === 0}
          className="flex items-center gap-1.5 px-3 py-2 bg-green-600 hover:bg-green-700 disabled:bg-gray-300 dark:disabled:bg-gray-700 disabled:cursor-not-allowed text-white text-xs font-semibold rounded-lg transition-colors"
        >
          <span className="material-symbols-outlined text-sm">check</span>
          {acceptedCount > 0 ? `Apply (${acceptedCount} accepted)` : 'Apply'}
        </button>
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
