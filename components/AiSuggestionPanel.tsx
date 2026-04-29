'use client';

import type { AiSuggestion } from '@/lib/types/ai-check';

interface AiSuggestionPanelProps {
  suggestion: AiSuggestion;
  onAccept: (improved: string) => void;
  onDismiss: () => void;
}

export default function AiSuggestionPanel({ suggestion, onAccept, onDismiss }: AiSuggestionPanelProps) {
  return (
    <div className="mx-8 mb-4 p-4 bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 rounded-lg">
      <div className="flex items-center gap-2 mb-3">
        <span className="material-symbols-outlined text-blue-600 dark:text-blue-400">auto_fix_high</span>
        <p className="text-sm font-bold text-blue-700 dark:text-blue-400">AI Suggestion</p>
        <span className="text-xs text-blue-500 dark:text-blue-500 italic ml-auto">Changes highlighted below</span>
      </div>

      {/* Diff view */}
      <div className="bg-white dark:bg-gray-900 rounded-lg p-4 text-base leading-relaxed font-display mb-4 border border-blue-100 dark:border-blue-900">
        {suggestion.diff.map((token, i) => {
          if (token.type === 'unchanged') {
            return <span key={i}>{token.text}</span>;
          }
          if (token.type === 'removed') {
            return (
              <span
                key={i}
                className="bg-red-100 dark:bg-red-900/40 text-red-700 dark:text-red-400 line-through rounded px-0.5"
              >
                {token.text}
              </span>
            );
          }
          // added
          return (
            <span
              key={i}
              className="bg-green-100 dark:bg-green-900/40 text-green-700 dark:text-green-400 underline rounded px-0.5"
            >
              {token.text}
            </span>
          );
        })}
      </div>

      {/* Legend */}
      <div className="flex items-center gap-4 text-xs text-gray-500 dark:text-gray-400 mb-4">
        <span className="flex items-center gap-1">
          <span className="inline-block w-3 h-3 rounded bg-red-100 dark:bg-red-900/40 border border-red-300"></span>
          Removed
        </span>
        <span className="flex items-center gap-1">
          <span className="inline-block w-3 h-3 rounded bg-green-100 dark:bg-green-900/40 border border-green-300"></span>
          Added
        </span>
      </div>

      {/* Actions */}
      <div className="flex gap-3">
        <button
          onClick={() => onAccept(suggestion.improved)}
          className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium rounded-lg transition-colors"
        >
          <span className="material-symbols-outlined text-base">check</span>
          Accept suggestion
        </button>
        <button
          onClick={onDismiss}
          className="flex items-center gap-2 px-4 py-2 bg-white dark:bg-gray-900 hover:bg-gray-50 dark:hover:bg-gray-800 text-gray-700 dark:text-gray-300 text-sm font-medium rounded-lg border border-gray-200 dark:border-gray-700 transition-colors"
        >
          <span className="material-symbols-outlined text-base">close</span>
          Dismiss
        </button>
      </div>
    </div>
  );
}
