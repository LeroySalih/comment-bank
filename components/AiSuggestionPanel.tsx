'use client';

import { useState, useMemo, useEffect, useRef } from 'react';
import type { AiSuggestion, DiffSegment } from '@/lib/types/ai-check';
import { groupChanges } from '@/lib/utils/diff';

interface AiSuggestionPanelProps {
  suggestion: AiSuggestion;
  title?: string;
  iconName?: string;
  maxGap?: number;
  showRerun?: boolean;
  onApplyChange: (newText: string, isLastChange: boolean) => void;
  onApplyAll: (improvedText: string) => void;
  onRejectChange?: (groupId: number, removedText: string) => void;
  onRejectAll?: (rejected: { id: number; removedText: string }[]) => void;
  onAllResolved?: () => void;
  onRerun?: () => void;
  onDismiss: () => void;
}

function humanizeKey(key: string): string {
  return key
    .replace(/_/g, ' ')
    .replace(/([A-Z])/g, ' $1')
    .trim();
}

function buildText(
  segments: DiffSegment[],
  applied: Map<number, string>
): string {
  return segments
    .map(seg => {
      if (seg.type === 'unchanged') return seg.tokens.map(t => t.text).join('');
      if (applied.has(seg.group.id)) return applied.get(seg.group.id) ?? '';
      return seg.group.removed.map(t => t.text).join('');
    })
    .join('');
}

export default function AiSuggestionPanel({
  suggestion,
  title = 'AI Suggestion',
  iconName = 'auto_fix_high',
  maxGap,
  showRerun = false,
  onApplyChange,
  onApplyAll,
  onRejectChange,
  onRejectAll,
  onAllResolved,
  onRerun,
  onDismiss,
}: AiSuggestionPanelProps) {
  const segments = useMemo(
    () => groupChanges(suggestion.diff, maxGap),
    [suggestion.diff, maxGap]
  );

  if (typeof window !== 'undefined') {
    console.log('[AiSuggestionPanel] render', {
      title,
      diffTokens: suggestion.diff.length,
      segments: segments.length,
      changeGroups: segments.filter(s => s.type === 'change').length,
    });
  }

  // Map of groupId → applied replacement text (default suggestion, custom edit, or '' for delete)
  const [applied, setApplied] = useState<Map<number, string>>(new Map());
  const [rejectedIds, setRejectedIds] = useState<Set<number>>(new Set());
  const [editingId, setEditingId] = useState<number | null>(null);
  const [editValue, setEditValue] = useState('');
  const editInputRef = useRef<HTMLInputElement | null>(null);

  const changeGroups = useMemo(
    () => segments.filter((s): s is Extract<DiffSegment, { type: 'change' }> => s.type === 'change'),
    [segments]
  );
  const totalChanges = changeGroups.length;
  const remainingGroups = changeGroups.filter(
    s => !applied.has(s.group.id) && !rejectedIds.has(s.group.id)
  );

  useEffect(() => {
    if (totalChanges > 0 && remainingGroups.length === 0 && onAllResolved) {
      onAllResolved();
    }
  }, [applied, rejectedIds, totalChanges, remainingGroups.length, onAllResolved]);

  useEffect(() => {
    if (editingId !== null && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [editingId]);

  const ruleEntries = Object.entries(suggestion.ruleChecks ?? {});
  const passedCount = ruleEntries.filter(([, v]) => v).length;

  const isLastResolution = (
    nextApplied: Map<number, string>,
    nextRejected: Set<number>
  ): boolean =>
    changeGroups.every(s => nextApplied.has(s.group.id) || nextRejected.has(s.group.id));

  const applyResolution = (groupId: number, replacement: string) => {
    const next = new Map(applied);
    next.set(groupId, replacement);
    setApplied(next);
    const newText = buildText(segments, next);
    onApplyChange(newText, isLastResolution(next, rejectedIds));
  };

  const handleApplyDefault = (group: Extract<DiffSegment, { type: 'change' }>['group']) => {
    const replacement = group.added.map(t => t.text).join('');
    applyResolution(group.id, replacement);
  };

  const handleDelete = (group: Extract<DiffSegment, { type: 'change' }>['group']) => {
    applyResolution(group.id, '');
  };

  const handleEditStart = (group: Extract<DiffSegment, { type: 'change' }>['group']) => {
    const initial =
      group.added.length > 0
        ? group.added.map(t => t.text).join('')
        : group.removed.map(t => t.text).join('');
    setEditValue(initial);
    setEditingId(group.id);
  };

  const handleEditSubmit = () => {
    if (editingId === null) return;
    applyResolution(editingId, editValue);
    setEditingId(null);
    setEditValue('');
  };

  const handleEditCancel = () => {
    setEditingId(null);
    setEditValue('');
  };

  const handleRejectChange = (id: number, removedText: string) => {
    const newRejected = new Set([...rejectedIds, id]);
    setRejectedIds(newRejected);
    onRejectChange?.(id, removedText);
  };

  const handleApplyAll = () => {
    const next = new Map(applied);
    for (const s of changeGroups) {
      if (next.has(s.group.id) || rejectedIds.has(s.group.id)) continue;
      if (s.group.added.length === 0) continue; // skip no-replacement entries
      next.set(s.group.id, s.group.added.map(t => t.text).join(''));
    }
    setApplied(next);
    onApplyAll(buildText(segments, next));
  };

  const handleRejectAll = () => {
    const toReject = changeGroups
      .filter(s => !applied.has(s.group.id) && !rejectedIds.has(s.group.id))
      .map(s => ({
        id: s.group.id,
        removedText: s.group.removed.map(t => t.text).join(''),
      }));
    const newRejected = new Set([...rejectedIds, ...toReject.map(r => r.id)]);
    setRejectedIds(newRejected);
    onRejectAll?.(toReject);
  };

  return (
    <div className="mx-8 mb-4 bg-white dark:bg-gray-900 border border-blue-200 dark:border-blue-800 rounded-lg overflow-hidden">

      {/* Header */}
      <div className="bg-blue-50 dark:bg-blue-900/20 border-b border-blue-200 dark:border-blue-800 px-4 py-3 flex items-center gap-2">
        <span className="material-symbols-outlined text-blue-600 dark:text-blue-400 text-base">{iconName}</span>
        <strong className="text-blue-700 dark:text-blue-400 text-sm">{title}</strong>
        <span className="ml-auto text-xs text-gray-500 dark:text-gray-400 italic">
          {remainingGroups.length > 0
            ? `${remainingGroups.length} of ${totalChanges} changes remaining`
            : 'All changes resolved'}
        </span>
        {showRerun && onRerun && (
          <button
            onClick={onRerun}
            className="ml-3 flex items-center gap-1 text-xs text-blue-600 dark:text-blue-400 hover:text-blue-700 dark:hover:text-blue-300 transition-colors"
            title="Re-run check"
          >
            <span className="material-symbols-outlined text-sm">refresh</span>
            Re-run
          </button>
        )}
      </div>

      {/* Diff section */}
      <div className="px-4 py-3 border-b border-gray-100 dark:border-gray-800">
        <div className="bg-gray-50 dark:bg-gray-950 border border-gray-200 dark:border-gray-700 rounded-lg p-3 text-base leading-loose font-display">
          {segments.map((seg, i) => {
            if (seg.type === 'unchanged') {
              return <span key={i}>{seg.tokens.map(t => t.text).join('')}</span>;
            }
            const { group } = seg;

            if (applied.has(group.id)) {
              const txt = applied.get(group.id) ?? '';
              if (txt === '') {
                return (
                  <span key={i} className="text-gray-400 dark:text-gray-500 italic" title="Deleted">
                    [deleted]
                  </span>
                );
              }
              return (
                <span key={i} className="text-green-700 dark:text-green-400">
                  {txt}
                </span>
              );
            }

            if (rejectedIds.has(group.id)) {
              return (
                <span key={i} className="text-gray-500 dark:text-gray-400" title="Rejected">
                  {group.removed.map(t => t.text).join('')}
                </span>
              );
            }

            if (editingId === group.id) {
              return (
                <span key={i} className="inline-flex items-center gap-1 mx-0.5 align-middle">
                  <input
                    ref={editInputRef}
                    type="text"
                    value={editValue}
                    onChange={e => setEditValue(e.target.value)}
                    onKeyDown={e => {
                      if (e.key === 'Enter') {
                        e.preventDefault();
                        handleEditSubmit();
                      } else if (e.key === 'Escape') {
                        e.preventDefault();
                        handleEditCancel();
                      }
                    }}
                    className="px-1.5 py-0.5 border border-blue-400 rounded text-sm bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-300"
                    style={{ minWidth: '6rem', width: `${Math.max(editValue.length + 2, 6)}ch` }}
                  />
                  <button
                    onClick={handleEditSubmit}
                    className="text-green-600 hover:text-green-700"
                    title="Apply"
                  >
                    <span className="material-symbols-outlined text-sm">check</span>
                  </button>
                  <button
                    onClick={handleEditCancel}
                    className="text-gray-400 hover:text-gray-600"
                    title="Cancel"
                  >
                    <span className="material-symbols-outlined text-sm">close</span>
                  </button>
                </span>
              );
            }

            const removedText = group.removed.map(t => t.text).join('');
            const addedText = group.added.map(t => t.text).join('');
            const canAccept = group.added.length > 0;

            return (
              <span key={i} className="inline-flex items-center gap-0.5 mx-0.5">
                <span
                  onClick={canAccept ? () => handleApplyDefault(group) : undefined}
                  className={canAccept ? 'cursor-pointer inline' : 'inline'}
                  title={canAccept ? 'Click to accept this change' : 'No suggestion — use edit, delete, or reject'}
                >
                  {group.removed.length > 0 && (
                    <span className="bg-red-50 dark:bg-red-900/30 text-red-600 dark:text-red-400 line-through rounded px-1 border border-red-200 dark:border-red-800 mx-0.5">
                      {removedText}
                    </span>
                  )}
                  {canAccept && group.removed.length > 0 && (
                    <span className="text-gray-400 text-xs mx-0.5">→</span>
                  )}
                  {canAccept && (
                    <span className="bg-green-50 dark:bg-green-900/30 text-green-700 dark:text-green-400 rounded px-1 border border-dashed border-green-300 dark:border-green-700 mx-0.5">
                      {addedText}
                    </span>
                  )}
                </span>
                <button
                  onClick={() => handleEditStart(group)}
                  className="ml-0.5 text-blue-500 hover:text-blue-700 dark:text-blue-400 dark:hover:text-blue-200"
                  title="Edit replacement"
                >
                  <span className="material-symbols-outlined text-xs">edit</span>
                </button>
                <button
                  onClick={() => handleDelete(group)}
                  className="ml-0.5 text-orange-500 hover:text-orange-700 dark:text-orange-400 dark:hover:text-orange-200"
                  title="Delete this word"
                >
                  <span className="material-symbols-outlined text-xs">delete</span>
                </button>
                {onRejectChange && (
                  <button
                    onClick={() => handleRejectChange(group.id, removedText)}
                    className="ml-0.5 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200"
                    title={canAccept ? 'Reject this change' : 'Ignore this word permanently'}
                  >
                    <span className="material-symbols-outlined text-xs">close</span>
                  </button>
                )}
              </span>
            );
          })}
        </div>
        {remainingGroups.length > 0 && (
          <p className="mt-2 text-xs text-gray-400 flex items-center gap-1">
            <span className="material-symbols-outlined text-sm">touch_app</span>
            Click a change to accept it, or use edit / delete / reject icons
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
          <>
            <button
              onClick={handleApplyAll}
              className="flex items-center gap-1.5 px-3 py-2 bg-green-600 hover:bg-green-700 text-white text-xs font-semibold rounded-lg transition-colors"
            >
              <span className="material-symbols-outlined text-sm">done_all</span>
              Accept all
            </button>
            {onRejectAll && (
              <button
                onClick={handleRejectAll}
                className="flex items-center gap-1.5 px-3 py-2 bg-gray-100 dark:bg-gray-800 hover:bg-gray-200 dark:hover:bg-gray-700 text-gray-700 dark:text-gray-300 text-xs font-semibold rounded-lg border border-gray-200 dark:border-gray-700 transition-colors"
              >
                <span className="material-symbols-outlined text-sm">close</span>
                Reject all
              </button>
            )}
          </>
        )}
        <button
          onClick={onDismiss}
          className="flex items-center gap-1.5 px-3 py-2 ml-auto bg-white dark:bg-gray-900 hover:bg-gray-50 dark:hover:bg-gray-800 text-gray-600 dark:text-gray-400 text-xs font-medium rounded-lg border border-gray-200 dark:border-gray-700 transition-colors"
        >
          <span className="material-symbols-outlined text-sm">close</span>
          Close
        </button>
      </div>
    </div>
  );
}
