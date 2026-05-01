'use client';

import { useEffect, useRef, useState } from 'react';
import type { StandardsResult, StandardsRuleKey } from '@/lib/types/ai-check';

// CSS keyframe + class for the slide-in animation injected once globally
const SIDEBAR_STYLE_ID = 'standards-panel-slide-in';
const SIDEBAR_KEYFRAMES = `
@keyframes standardsPanelSlideIn {
  from { transform: translateX(100%); }
  to   { transform: translateX(0); }
}
.standards-panel-slide-in { animation: standardsPanelSlideIn 220ms ease-out; }
`;

interface StandardsPanelProps {
  result: StandardsResult;
  onAllResolved?: () => void;
  onRerun: () => void;
  onDismiss: () => void;
}

type Rule = {
  key: StandardsRuleKey;
  label: string;
  description: string;
};

const RULES: Rule[] = [
  {
    key: 'UKSpelling',
    label: 'UK Spelling',
    description: 'All text in the comment field must use British English spelling.',
  },
  {
    key: 'CourseOverviewIncluded',
    label: 'Course Overview Included',
    description: 'Ensure that the comment includes an overview of the units taken by the pupil this year.',
  },
  {
    key: 'AcademicPerformanceIncluded',
    label: 'Academic Performance Included',
    description: 'Ensure that the comment includes the academic performance of the pupil.',
  },
  {
    key: 'TargetWordCountMet',
    label: 'Target Word Count Met',
    description: 'Ensure that the number of words of the comment is between 250 and 280.',
  },
  {
    key: 'TerminologyCorrect',
    label: 'Terminology Correct',
    description: 'Ensure that the terminology used in the comment is correct to the context. Ensure that the word "pupil" is used instead of "teacher".',
  },
  {
    key: 'JargonFree',
    label: 'Jargon Free',
    description: 'Ensure that the comment does not include any jargon.',
  },
  {
    key: 'DataSpecific',
    label: 'Data Specific',
    description: 'Ensure that the performance of the pupil includes specific data values, such as 2M or 7H.',
  },
  {
    key: 'ToneBalanced',
    label: 'Tone Balanced',
    description: 'Ensure that the tone of the comment is balanced and professional.',
  },
  {
    key: 'SocialSkillsIncluded',
    label: 'Social Skills Included',
    description: 'Ensure that the comment includes how the pupil interacts with other pupils.',
  },
  {
    key: 'CollaborationIncluded',
    label: 'Collaboration Included',
    description: 'Ensure that the comment includes a section on the collaboration displayed by the pupil.',
  },
  {
    key: 'BehaviourIncluded',
    label: 'Behaviour Included',
    description: 'Ensure that the comment includes a section on the behaviour of the pupil.',
  },
  {
    key: 'ParentalSupportIncluded',
    label: 'Parental Support Included',
    description: 'Ensure that the comment includes a description of how the pupil can improve in the future.',
  },
  {
    key: 'Formatting',
    label: 'Formatting',
    description: 'Ensure a clear space between paragraphs for readability. End of sentences should have a double space.',
  },
];

export default function StandardsPanel({ result, onAllResolved, onRerun, onDismiss }: StandardsPanelProps) {
  const [openKey, setOpenKey] = useState<StandardsRuleKey | null>(null);
  const popoverRefs = useRef<Map<string, HTMLDivElement | null>>(new Map());

  const passedCount = RULES.filter(r => result[r.key]?.result === true).length;
  const totalCount = RULES.length;
  const allPassed = result.Status?.result === true && passedCount === totalCount;

  useEffect(() => {
    if (allPassed && onAllResolved) onAllResolved();
  }, [allPassed, onAllResolved]);

  // Inject animation keyframes once
  useEffect(() => {
    if (document.getElementById(SIDEBAR_STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = SIDEBAR_STYLE_ID;
    style.textContent = SIDEBAR_KEYFRAMES;
    document.head.appendChild(style);
  }, []);

  // Close popover on outside click
  useEffect(() => {
    if (openKey === null) return;
    const handler = (e: MouseEvent) => {
      const popover = popoverRefs.current.get(openKey);
      if (popover && !popover.contains(e.target as Node)) {
        // Don't close if clicking the rule row itself (handled by toggle)
        const row = document.querySelector(`[data-rule-key="${openKey}"]`);
        if (row && row.contains(e.target as Node)) return;
        setOpenKey(null);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [openKey]);

  return (
    <aside
      className="standards-panel-slide-in fixed right-0 top-0 bottom-0 z-[60] w-full sm:w-[420px] bg-white dark:bg-gray-900 border-l border-blue-200 dark:border-blue-800 shadow-2xl flex flex-col"
      role="complementary"
      aria-label="Standards check results"
    >
      {/* Header */}
      <div className="bg-blue-50 dark:bg-blue-900/20 border-b border-blue-200 dark:border-blue-800 px-3 py-3 flex items-center gap-2 flex-shrink-0">
        <button
          onClick={onDismiss}
          className="flex items-center justify-center w-8 h-8 rounded-lg bg-white dark:bg-gray-800 border border-gray-300 dark:border-gray-600 text-gray-600 dark:text-gray-300 hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors flex-shrink-0"
          title="Close sidebar"
          aria-label="Close sidebar"
        >
          <span className="material-symbols-outlined text-lg">close</span>
        </button>
        <span className="material-symbols-outlined text-blue-600 dark:text-blue-400 text-base ml-1">checklist</span>
        <strong className="text-blue-700 dark:text-blue-400 text-sm">Standards Check</strong>
        <button
          onClick={onRerun}
          className="ml-auto flex items-center gap-1 px-2 py-1.5 rounded text-xs font-medium text-blue-700 dark:text-blue-400 bg-white dark:bg-gray-800 border border-blue-200 dark:border-blue-800 hover:bg-blue-100 dark:hover:bg-blue-900/40 transition-colors"
          title="Re-run standards check"
        >
          <span className="material-symbols-outlined text-sm">refresh</span>
          Re-run
        </button>
      </div>

      {/* Subtitle */}
      <div className="px-4 py-2 text-xs text-gray-500 dark:text-gray-400 italic border-b border-gray-100 dark:border-gray-800 flex-shrink-0">
        {allPassed
          ? 'All standards met'
          : `${passedCount} of ${totalCount} standards met — fix the failures and re-run`}
      </div>

      {/* Scrollable body */}
      <div className="flex-1 overflow-y-auto">

      {/* Status banner */}
      <div className={`px-4 py-3 border-b text-sm flex items-start gap-2 ${
        allPassed
          ? 'bg-green-50 dark:bg-green-900/20 border-green-200 dark:border-green-800 text-green-800 dark:text-green-300'
          : 'bg-red-50 dark:bg-red-900/20 border-red-200 dark:border-red-800 text-red-800 dark:text-red-300'
      }`}>
        <span className="material-symbols-outlined text-base flex-shrink-0">
          {allPassed ? 'check_circle' : 'error'}
        </span>
        <div className="flex-1 min-w-0">
          <strong className="block">{allPassed ? 'Passed' : 'Failed'}</strong>
          <span className="text-xs opacity-80">
            {allPassed
              ? 'You can now copy the comment to the clipboard.'
              : 'Edit the comment to address the failed rules, then re-run the check.'}
          </span>
        </div>
      </div>

      {/* Rule list */}
      <div className="px-4 py-3">
        <div className="grid grid-cols-1 gap-2">
          {RULES.map(rule => {
            const entry = result[rule.key];
            const passed = entry?.result === true;
            const hasInstances = entry?.instances && entry.instances.length > 0;
            const hasWordCount = typeof entry?.wordCount === 'number';
            const isOpen = openKey === rule.key;
            return (
              <div key={rule.key} className="relative">
                <button
                  data-rule-key={rule.key}
                  onClick={() => setOpenKey(isOpen ? null : rule.key)}
                  className={`w-full text-left flex items-center gap-2 px-3 py-2 rounded text-sm transition-colors ${
                    passed
                      ? 'bg-green-50 dark:bg-green-900/20 text-green-800 dark:text-green-300 hover:bg-green-100 dark:hover:bg-green-900/40'
                      : 'bg-red-50 dark:bg-red-900/20 text-red-800 dark:text-red-300 hover:bg-red-100 dark:hover:bg-red-900/40 border-l-2 border-red-400'
                  }`}
                >
                  <span className="material-symbols-outlined text-base flex-shrink-0">
                    {passed ? 'check_circle' : 'cancel'}
                  </span>
                  <span className="flex-1 truncate">{rule.label}</span>
                  {hasWordCount && (
                    <span className="text-[11px] font-mono px-1.5 py-0.5 rounded bg-white/60 dark:bg-gray-900/60 flex-shrink-0">
                      {entry!.wordCount} words
                    </span>
                  )}
                  {hasInstances && (
                    <span
                      className="text-[11px] font-mono px-1.5 py-0.5 rounded bg-white/60 dark:bg-gray-900/60 flex-shrink-0"
                      title={`${entry!.instances!.length} flagged ${entry!.instances!.length === 1 ? 'instance' : 'instances'}`}
                    >
                      {entry!.instances!.length}×
                    </span>
                  )}
                  <span className="material-symbols-outlined text-sm opacity-50 flex-shrink-0">
                    info
                  </span>
                </button>
                {isOpen && (
                  <div
                    ref={el => {
                      popoverRefs.current.set(rule.key, el);
                    }}
                    className="absolute z-50 left-0 right-0 top-full mt-1 bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-lg shadow-lg p-3 text-sm"
                  >
                    <div className="flex items-start gap-2">
                      <span className={`material-symbols-outlined text-base flex-shrink-0 ${passed ? 'text-green-600 dark:text-green-400' : 'text-red-600 dark:text-red-400'}`}>
                        {passed ? 'check_circle' : 'cancel'}
                      </span>
                      <div className="flex-1 min-w-0">
                        <p className="font-semibold text-gray-900 dark:text-gray-100">{rule.label}</p>
                        <p className="text-xs text-gray-600 dark:text-gray-400 mt-1">{rule.description}</p>
                        <p className={`text-[11px] uppercase font-semibold mt-2 ${passed ? 'text-green-600 dark:text-green-400' : 'text-red-600 dark:text-red-400'}`}>
                          {passed ? 'PASSED' : 'FAILED'}
                        </p>

                        {hasWordCount && (
                          <div className="mt-2 pt-2 border-t border-gray-100 dark:border-gray-800">
                            <p className="text-[10px] uppercase font-semibold text-gray-400 mb-1">Word count</p>
                            <p className="text-sm font-mono text-gray-700 dark:text-gray-300">
                              {entry!.wordCount}
                              <span className="text-xs text-gray-400 ml-2">(target 250–280)</span>
                            </p>
                          </div>
                        )}

                        {hasInstances && (
                          <div className="mt-2 pt-2 border-t border-gray-100 dark:border-gray-800">
                            <p className="text-[10px] uppercase font-semibold text-gray-400 mb-1">
                              Flagged {entry!.instances!.length === 1 ? 'instance' : 'instances'}
                            </p>
                            <ul className="space-y-1 max-h-40 overflow-y-auto">
                              {entry!.instances!.map((inst, idx) => (
                                <li
                                  key={idx}
                                  className="text-xs px-2 py-1 bg-red-50 dark:bg-red-900/20 text-red-700 dark:text-red-400 border border-red-200 dark:border-red-800 rounded font-mono break-words"
                                >
                                  {inst}
                                </li>
                              ))}
                            </ul>
                          </div>
                        )}
                      </div>
                      <button
                        onClick={() => setOpenKey(null)}
                        className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-200"
                        title="Close"
                      >
                        <span className="material-symbols-outlined text-sm">close</span>
                      </button>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      </div>{/* /scrollable body */}
    </aside>
  );
}
