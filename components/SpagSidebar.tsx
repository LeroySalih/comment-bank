'use client';

import { useEffect } from 'react';
import type { SpagMatch } from '@/lib/types/ai-check';
import type { SpagResolution } from './SpagPanel';

const SIDEBAR_STYLE_ID = 'spag-sidebar-slide-in';
const SIDEBAR_KEYFRAMES = `
@keyframes spagSidebarSlideIn {
  from { transform: translateX(100%); }
  to   { transform: translateX(0); }
}
.spag-sidebar-slide-in { animation: spagSidebarSlideIn 220ms ease-out; }
`;

interface SpagSidebarProps {
  matches: SpagMatch[];
  resolutions: Map<number, SpagResolution>;
  onJumpTo: (match: SpagMatch) => void;
  onAcceptAll: () => void;
  onRejectAll: () => void;
  onRerun: () => void;
  onDismiss: () => void;
}

export default function SpagSidebar({
  matches,
  resolutions,
  onJumpTo,
  onAcceptAll,
  onRejectAll,
  onRerun,
  onDismiss,
}: SpagSidebarProps) {
  useEffect(() => {
    if (document.getElementById(SIDEBAR_STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = SIDEBAR_STYLE_ID;
    style.textContent = SIDEBAR_KEYFRAMES;
    document.head.appendChild(style);
  }, []);

  const sortedMatches = [...matches].sort((a, b) => a.offset - b.offset);
  const totalCount = matches.length;
  const resolvedCount = resolutions.size;
  const unresolvedCount = totalCount - resolvedCount;

  const renderStatusBadge = (resolution: SpagResolution | undefined) => {
    if (!resolution) {
      return (
        <span className="text-[10px] uppercase font-semibold px-1.5 py-0.5 rounded bg-amber-100 dark:bg-amber-900/40 text-amber-700 dark:text-amber-400 flex-shrink-0">
          Pending
        </span>
      );
    }
    if (resolution.action === 'accepted') {
      return (
        <span className="text-[10px] uppercase font-semibold px-1.5 py-0.5 rounded bg-green-100 dark:bg-green-900/40 text-green-700 dark:text-green-400 flex-shrink-0">
          Accepted
        </span>
      );
    }
    if (resolution.action === 'edited') {
      return (
        <span className="text-[10px] uppercase font-semibold px-1.5 py-0.5 rounded bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-400 flex-shrink-0">
          Edited
        </span>
      );
    }
    if (resolution.action === 'deleted') {
      return (
        <span className="text-[10px] uppercase font-semibold px-1.5 py-0.5 rounded bg-orange-100 dark:bg-orange-900/40 text-orange-700 dark:text-orange-400 flex-shrink-0">
          Deleted
        </span>
      );
    }
    // ignored
    return (
      <span className="text-[10px] uppercase font-semibold px-1.5 py-0.5 rounded bg-gray-200 dark:bg-gray-800 text-gray-600 dark:text-gray-400 flex-shrink-0">
        Ignored
      </span>
    );
  };

  return (
    <aside
      className="spag-sidebar-slide-in fixed right-0 top-0 bottom-0 z-[60] w-full sm:w-[420px] bg-white dark:bg-gray-900 border-l border-amber-200 dark:border-amber-800 shadow-2xl flex flex-col"
      role="complementary"
      aria-label="Spelling and grammar errors"
    >
      {/* Header */}
      <div className="bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-800 px-3 py-3 flex items-center gap-2 flex-shrink-0">
        <button
          onClick={onDismiss}
          className="flex items-center justify-center w-8 h-8 rounded-lg bg-white dark:bg-gray-800 border border-gray-300 dark:border-gray-600 text-gray-600 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors flex-shrink-0"
          title="Close sidebar"
          aria-label="Close sidebar"
        >
          <span className="material-symbols-outlined text-lg">close</span>
        </button>
        <span className="material-symbols-outlined text-amber-600 dark:text-amber-400 text-base ml-1">spellcheck</span>
        <strong className="text-amber-700 dark:text-amber-400 text-sm">SPAG Errors</strong>
        <button
          onClick={onRerun}
          className="ml-auto flex items-center gap-1 px-2 py-1.5 rounded text-xs font-medium text-amber-700 dark:text-amber-400 bg-white dark:bg-gray-800 border border-amber-200 dark:border-amber-800 hover:bg-amber-100 dark:hover:bg-amber-900/40 transition-colors"
          title="Re-run SPAG check"
        >
          <span className="material-symbols-outlined text-sm">refresh</span>
          Re-run
        </button>
      </div>

      {/* Subtitle */}
      <div className="px-4 py-2 text-xs text-gray-500 dark:text-gray-400 italic border-b border-gray-100 dark:border-gray-800 flex-shrink-0">
        {totalCount === 0
          ? 'No errors found'
          : unresolvedCount === 0
            ? `All ${totalCount} resolved`
            : `${resolvedCount} of ${totalCount} resolved — ${unresolvedCount} pending`}
      </div>

      {/* Match list (scrollable) */}
      <div className="flex-1 overflow-y-auto">
        {sortedMatches.length === 0 ? (
          <div className="px-4 py-6 text-center text-sm text-gray-500 dark:text-gray-400">
            No spelling or grammar issues found.
          </div>
        ) : (
          <ul className="divide-y divide-gray-100 dark:divide-gray-800">
            {sortedMatches.map((match, idx) => {
              const resolution = resolutions.get(match.offset);
              const replacementText =
                resolution && (resolution.action === 'accepted' || resolution.action === 'edited')
                  ? resolution.replacement
                  : null;

              return (
                <li key={`${match.offset}-${idx}`}>
                  <button
                    onClick={() => onJumpTo(match)}
                    className="w-full text-left px-4 py-3 hover:bg-amber-50 dark:hover:bg-amber-900/10 transition-colors flex items-start gap-2"
                    title="Jump to this word in the comment"
                  >
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="font-mono font-semibold text-red-700 dark:text-red-400 text-sm">
                          {match.word}
                        </span>
                        {replacementText !== null && (
                          <>
                            <span className="text-xs text-gray-400">→</span>
                            <span className="font-mono text-green-700 dark:text-green-400 text-sm">
                              {replacementText === '' ? '(deleted)' : replacementText}
                            </span>
                          </>
                        )}
                        {renderStatusBadge(resolution)}
                      </div>
                      <p className="text-xs text-gray-500 dark:text-gray-400 mt-1 leading-snug line-clamp-2">
                        {match.message}
                      </p>
                    </div>
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </div>

      {/* Bulk actions footer */}
      {unresolvedCount > 0 && (
        <div className="border-t border-gray-100 dark:border-gray-800 px-4 py-3 flex items-center gap-2 flex-shrink-0">
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
            title="Ignore every remaining issue"
          >
            <span className="material-symbols-outlined text-sm">visibility_off</span>
            Reject all
          </button>
        </div>
      )}
    </aside>
  );
}
