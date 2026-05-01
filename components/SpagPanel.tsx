'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import type { SpagMatch } from '@/lib/types/ai-check';

export type SpagResolution =
  | { match: SpagMatch; action: 'accepted'; replacement: string }
  | { match: SpagMatch; action: 'edited'; replacement: string }
  | { match: SpagMatch; action: 'deleted' }
  | { match: SpagMatch; action: 'ignored' };

interface SpagPanelProps {
  originalText: string;
  matches: SpagMatch[];
  resolutions: Map<number, SpagResolution>;
  onResolve: (resolution: SpagResolution) => void;
  onAcceptAll: () => void;
  onRejectAll: () => void;
  onAllResolved?: () => void;
  onRerun: () => void;
  onDismiss: () => void;
}

type Segment =
  | { type: 'unchanged'; text: string }
  | { type: 'match'; match: SpagMatch };

export default function SpagPanel({
  originalText,
  matches,
  resolutions,
  onResolve,
  onAcceptAll,
  onRejectAll,
  onAllResolved,
  onRerun,
  onDismiss,
}: SpagPanelProps) {
  const [activeOffset, setActiveOffset] = useState<number | null>(null);
  const [popupPlacement, setPopupPlacement] = useState<'below' | 'above'>('below');
  const [editingOffset, setEditingOffset] = useState<number | null>(null);
  const [editValue, setEditValue] = useState('');
  const containerRef = useRef<HTMLDivElement | null>(null);
  const popupRef = useRef<HTMLDivElement | null>(null);
  const editInputRef = useRef<HTMLInputElement | null>(null);

  const segments = useMemo<Segment[]>(() => {
    const sorted = [...matches].sort((a, b) => a.offset - b.offset);
    const out: Segment[] = [];
    let cursor = 0;
    for (const m of sorted) {
      if (m.offset > cursor) {
        out.push({ type: 'unchanged', text: originalText.slice(cursor, m.offset) });
      }
      out.push({ type: 'match', match: m });
      cursor = m.offset + m.length;
    }
    if (cursor < originalText.length) {
      out.push({ type: 'unchanged', text: originalText.slice(cursor) });
    }
    return out;
  }, [matches, originalText]);

  const totalCount = matches.length;
  const resolvedCount = resolutions.size;
  const unresolvedCount = totalCount - resolvedCount;
  const allDone = totalCount > 0 && unresolvedCount === 0;

  useEffect(() => {
    if (allDone && onAllResolved) onAllResolved();
  }, [allDone, onAllResolved]);

  // Click outside closes the popup
  useEffect(() => {
    if (activeOffset === null) return;
    const handler = (e: MouseEvent) => {
      if (popupRef.current && !popupRef.current.contains(e.target as Node)) {
        // Also check the underline span isn't being re-clicked
        const underline = containerRef.current?.querySelector(`[data-spag-offset="${activeOffset}"]`);
        if (underline && underline.contains(e.target as Node)) return;
        setActiveOffset(null);
        setEditingOffset(null);
        setEditValue('');
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [activeOffset]);

  useEffect(() => {
    if (editingOffset !== null && editInputRef.current) {
      editInputRef.current.focus();
      editInputRef.current.select();
    }
  }, [editingOffset]);

  const closePopup = () => {
    setActiveOffset(null);
    setEditingOffset(null);
    setEditValue('');
  };

  const handleOpenPopup = (e: React.MouseEvent<HTMLSpanElement>, match: SpagMatch) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom;
    const spaceAbove = rect.top;
    // Approximate popup height: ~220px when fully expanded
    const POPUP_HEIGHT = 240;
    setPopupPlacement(spaceBelow < POPUP_HEIGHT && spaceAbove > spaceBelow ? 'above' : 'below');
    setActiveOffset(match.offset);
    setEditingOffset(null);
    setEditValue('');
  };

  const handleAccept = (match: SpagMatch, replacement: string) => {
    onResolve({ match, action: 'accepted', replacement });
    closePopup();
  };

  const handleEditStart = (match: SpagMatch) => {
    setEditingOffset(match.offset);
    setEditValue(match.replacements[0] ?? match.word);
  };

  const handleEditSubmit = (match: SpagMatch) => {
    onResolve({ match, action: 'edited', replacement: editValue });
    closePopup();
  };

  const handleDelete = (match: SpagMatch) => {
    onResolve({ match, action: 'deleted' });
    closePopup();
  };

  const handleIgnore = (match: SpagMatch) => {
    onResolve({ match, action: 'ignored' });
    closePopup();
  };

  const renderResolved = (resolution: SpagResolution, key: string) => {
    if (resolution.action === 'accepted' || resolution.action === 'edited') {
      return (
        <span key={key} className="text-green-700 dark:text-green-400">
          {resolution.replacement}
        </span>
      );
    }
    if (resolution.action === 'deleted') {
      return (
        <span key={key} className="text-gray-400 dark:text-gray-500 italic" title="Deleted">
          [deleted]
        </span>
      );
    }
    // ignored
    return (
      <span key={key} className="text-gray-700 dark:text-gray-300" title="Ignored">
        {resolution.match.word}
      </span>
    );
  };

  return (
    <div className="mx-8 mb-4 bg-white dark:bg-gray-900 border border-amber-200 dark:border-amber-800 rounded-lg">

      {/* Header */}
      <div className="bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-800 px-4 py-3 flex items-center gap-2 rounded-t-lg">
        <span className="material-symbols-outlined text-amber-600 dark:text-amber-400 text-base">spellcheck</span>
        <strong className="text-amber-700 dark:text-amber-400 text-sm">SPAG Check</strong>
        <span className="ml-auto text-xs text-gray-500 dark:text-gray-400 italic">
          {unresolvedCount > 0
            ? `${unresolvedCount} of ${totalCount} unresolved — click any underlined word`
            : 'All issues resolved'}
        </span>
        <button
          onClick={onRerun}
          className="ml-3 flex items-center gap-1 text-xs text-amber-600 dark:text-amber-400 hover:text-amber-700 dark:hover:text-amber-300 transition-colors"
          title="Re-run SPAG check"
        >
          <span className="material-symbols-outlined text-sm">refresh</span>
          Re-run
        </button>
      </div>

      {/* Annotated text */}
      <div ref={containerRef} className="px-4 py-3 border-b border-gray-100 dark:border-gray-800">
        <div className="bg-gray-50 dark:bg-gray-950 border border-gray-200 dark:border-gray-700 rounded-lg p-3 text-base leading-loose font-display whitespace-pre-wrap">
          {segments.map((seg, i) => {
            if (seg.type === 'unchanged') {
              return <span key={i}>{seg.text}</span>;
            }
            const { match } = seg;
            const resolution = resolutions.get(match.offset);
            if (resolution) {
              return renderResolved(resolution, `${i}`);
            }

            // Pending — render with red wavy underline + click handler
            const isActive = activeOffset === match.offset;
            return (
              <span key={i} className="relative inline-block">
                <span
                  data-spag-offset={match.offset}
                  onClick={(e) => handleOpenPopup(e, match)}
                  className="cursor-pointer text-red-700 dark:text-red-400"
                  style={{
                    textDecorationLine: 'underline',
                    textDecorationStyle: 'wavy',
                    textDecorationColor: '#ef4444',
                    textDecorationThickness: '1.5px',
                    textUnderlineOffset: '3px',
                  }}
                  title={match.message}
                >
                  {match.word}
                </span>

                {isActive && (
                  <div
                    ref={popupRef}
                    className={`absolute z-50 left-0 ${popupPlacement === 'below' ? 'top-full mt-1' : 'bottom-full mb-1'} min-w-[14rem] max-w-[20rem] bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-lg shadow-lg p-2 text-sm font-sans whitespace-normal`}
                  >
                    <div className="flex items-start gap-2 px-1 pb-2 border-b border-gray-100 dark:border-gray-800">
                      <span className="material-symbols-outlined text-amber-500 text-base">spellcheck</span>
                      <div className="flex-1 min-w-0">
                        <p className="font-mono font-semibold text-red-600 dark:text-red-400 text-sm truncate">{match.word}</p>
                        <p className="text-xs text-gray-500 dark:text-gray-400 leading-snug">{match.message}</p>
                      </div>
                      <button
                        onClick={closePopup}
                        className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-200"
                        title="Close"
                      >
                        <span className="material-symbols-outlined text-sm">close</span>
                      </button>
                    </div>

                    {/* Edit mode */}
                    {editingOffset === match.offset ? (
                      <div className="pt-2 px-1 flex items-center gap-1">
                        <input
                          ref={editInputRef}
                          type="text"
                          value={editValue}
                          onChange={e => setEditValue(e.target.value)}
                          onKeyDown={e => {
                            if (e.key === 'Enter') {
                              e.preventDefault();
                              handleEditSubmit(match);
                            } else if (e.key === 'Escape') {
                              e.preventDefault();
                              closePopup();
                            }
                          }}
                          className="flex-1 px-2 py-1 border border-blue-400 rounded text-sm bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-300"
                        />
                        <button
                          onClick={() => handleEditSubmit(match)}
                          className="px-2 py-1 bg-green-600 hover:bg-green-700 text-white rounded text-xs font-semibold"
                        >
                          Apply
                        </button>
                      </div>
                    ) : (
                      <>
                        {/* Replacement options */}
                        {match.replacements.length > 0 ? (
                          <div className="pt-2 px-1">
                            <p className="text-[10px] uppercase font-semibold text-gray-400 mb-1">Suggestions</p>
                            <div className="flex flex-wrap gap-1">
                              {match.replacements.slice(0, 6).map(r => (
                                <button
                                  key={r}
                                  onClick={() => handleAccept(match, r)}
                                  className="px-2 py-1 bg-green-50 dark:bg-green-900/30 text-green-700 dark:text-green-400 border border-green-200 dark:border-green-700 rounded text-xs font-mono hover:bg-green-100 dark:hover:bg-green-900/50 transition-colors"
                                >
                                  {r}
                                </button>
                              ))}
                            </div>
                          </div>
                        ) : (
                          <p className="pt-2 px-1 text-xs text-gray-400 italic">No suggestions available</p>
                        )}

                        {/* Action buttons */}
                        <div className="pt-2 px-1 mt-2 border-t border-gray-100 dark:border-gray-800 flex items-center gap-1">
                          <button
                            onClick={() => handleEditStart(match)}
                            className="flex items-center gap-1 px-2 py-1 text-xs text-blue-600 dark:text-blue-400 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded"
                            title="Type a custom replacement"
                          >
                            <span className="material-symbols-outlined text-sm">edit</span>
                            Edit
                          </button>
                          <button
                            onClick={() => handleDelete(match)}
                            className="flex items-center gap-1 px-2 py-1 text-xs text-orange-600 dark:text-orange-400 hover:bg-orange-50 dark:hover:bg-orange-900/20 rounded"
                            title="Remove this word"
                          >
                            <span className="material-symbols-outlined text-sm">delete</span>
                            Delete
                          </button>
                          <button
                            onClick={() => handleIgnore(match)}
                            className="flex items-center gap-1 px-2 py-1 text-xs text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-800 rounded ml-auto"
                            title="Ignore this word permanently"
                          >
                            <span className="material-symbols-outlined text-sm">visibility_off</span>
                            Ignore
                          </button>
                        </div>
                      </>
                    )}
                  </div>
                )}
              </span>
            );
          })}
        </div>
      </div>

      {/* Bulk actions */}
      <div className="px-4 py-3 flex items-center gap-2">
        {unresolvedCount > 0 && (
          <>
            <button
              onClick={onAcceptAll}
              className="flex items-center gap-1.5 px-3 py-2 bg-green-600 hover:bg-green-700 text-white text-xs font-semibold rounded-lg transition-colors"
              title="Accept the first replacement for every remaining issue"
            >
              <span className="material-symbols-outlined text-sm">done_all</span>
              Accept all
            </button>
            <button
              onClick={onRejectAll}
              className="flex items-center gap-1.5 px-3 py-2 bg-gray-100 dark:bg-gray-800 hover:bg-gray-200 dark:hover:bg-gray-700 text-gray-700 dark:text-gray-300 text-xs font-semibold rounded-lg border border-gray-200 dark:border-gray-700 transition-colors"
              title="Ignore every remaining issue (words will not be flagged again)"
            >
              <span className="material-symbols-outlined text-sm">visibility_off</span>
              Reject all
            </button>
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
