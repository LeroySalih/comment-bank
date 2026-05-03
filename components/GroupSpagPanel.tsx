'use client'

import { useState, useEffect } from 'react'
import { requestSpagCheck } from '@/lib/server-actions/ai-check'
import { addIgnoredWord } from '@/lib/server-actions/ignored-words'
import { substituteVariables } from '@/lib/audit/substitute-variables'
import type { SpagMatch } from '@/lib/types/ai-check'

interface Option {
  id: string
  code: string
  text: string
}

type ItemResult =
  | { status: 'pending' }
  | { status: 'checking' }
  | { status: 'pass' }
  | { status: 'issues'; matches: SpagMatch[]; ignoredWords: Set<string> }
  | { status: 'error'; message: string }

interface GroupSpagPanelProps {
  options: Option[]
  subjectTitle?: string
  onClose: () => void
}

export function GroupSpagPanel({ options, subjectTitle, onClose }: GroupSpagPanelProps) {
  const [results, setResults] = useState<Map<string, ItemResult>>(
    () => new Map(options.map(o => [o.id, { status: 'pending' }]))
  )
  const [expandedItems, setExpandedItems] = useState<Set<string>>(new Set())
  const [isRunning, setIsRunning] = useState(false)

  // Start checking automatically when the panel mounts
  useEffect(() => { runChecks() }, []) // eslint-disable-line react-hooks/exhaustive-deps

  async function runChecks() {
    setIsRunning(true)
    setExpandedItems(new Set())
    setResults(new Map(options.map(o => [o.id, { status: 'pending' }])))

    for (const option of options) {
      setResults(prev => new Map(prev).set(option.id, { status: 'checking' }))

      const substituted = substituteVariables(option.text, subjectTitle || 'Computer Science')
      const result = await requestSpagCheck(substituted)

      if (result.success) {
        const matches = result.result?.matches ?? []
        setResults(prev => new Map(prev).set(option.id,
          matches.length === 0
            ? { status: 'pass' }
            : { status: 'issues', matches, ignoredWords: new Set() }
        ))
        if (matches.length > 0) {
          setExpandedItems(prev => new Set(prev).add(option.id))
        }
      } else {
        setResults(prev => new Map(prev).set(option.id,
          { status: 'error', message: result.error || 'Check failed' }
        ))
      }
    }

    setIsRunning(false)
  }

  async function handleIgnoreWord(optionId: string, word: string) {
    const result = await addIgnoredWord(word)
    if (!result.success) return
    setResults(prev => {
      const entry = prev.get(optionId)
      if (entry?.status !== 'issues') return prev
      const next = new Map(prev)
      next.set(optionId, { ...entry, ignoredWords: new Set(entry.ignoredWords).add(word.toLowerCase()) })
      return next
    })
  }

  function toggleItem(id: string) {
    setExpandedItems(prev => {
      const next = new Set(prev)
      next.has(id) ? next.delete(id) : next.add(id)
      return next
    })
  }

  const total = options.length
  const done = [...results.values()].filter(r => r.status !== 'pending' && r.status !== 'checking').length
  const issueItems = [...results.values()].filter(r => r.status === 'issues')
  const totalIssues = issueItems.reduce((sum, r) =>
    r.status === 'issues' ? sum + r.matches.filter(m => !r.ignoredWords.has(m.word.toLowerCase())).length : sum
  , 0)

  return (
    <div className="mb-4 border border-amber-200 dark:border-amber-800 rounded-lg overflow-hidden bg-white dark:bg-gray-900">
      {/* Header */}
      <div className="flex items-center gap-2 px-3 py-2 bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-800">
        <span className="material-symbols-outlined text-amber-600 dark:text-amber-400 text-base">spellcheck</span>
        <span className="text-sm font-semibold text-amber-700 dark:text-amber-400">
          {isRunning
            ? `Checking… ${done} of ${total}`
            : totalIssues > 0
            ? `${totalIssues} issue${totalIssues > 1 ? 's' : ''} across ${issueItems.length} item${issueItems.length > 1 ? 's' : ''}`
            : `All ${total} items passed ✓`}
        </span>
        <div className="ml-auto flex items-center gap-2">
          {!isRunning && (
            <button
              type="button"
              onClick={runChecks}
              className="flex items-center gap-1 text-xs text-amber-600 dark:text-amber-400 hover:text-amber-700 dark:hover:text-amber-300 transition-colors"
            >
              <span className="material-symbols-outlined text-sm">refresh</span>Re-run
            </button>
          )}
          <button type="button" onClick={onClose} className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-200">
            <span className="material-symbols-outlined text-sm">close</span>
          </button>
        </div>
      </div>

      {/* Per-item results */}
      <ul className="divide-y divide-gray-100 dark:divide-gray-800">
        {options.map(option => {
          const result = results.get(option.id) ?? { status: 'pending' as const }
          const isExpanded = expandedItems.has(option.id)
          const visibleMatches = result.status === 'issues'
            ? result.matches.filter(m => !result.ignoredWords.has(m.word.toLowerCase()))
            : []

          return (
            <li key={option.id}>
              <div
                className={`flex items-center gap-2 px-3 py-2 ${result.status === 'issues' ? 'cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-800/50' : ''}`}
                onClick={() => result.status === 'issues' && toggleItem(option.id)}
              >
                <span className="text-xs font-mono font-bold text-gray-500 dark:text-gray-400 w-8 shrink-0">
                  {option.code}
                </span>
                <span className="text-xs text-gray-700 dark:text-gray-300 flex-1 truncate">
                  {option.text.length > 70 ? option.text.slice(0, 70) + '…' : option.text}
                </span>

                {result.status === 'pending' && (
                  <span className="text-[10px] text-gray-400 shrink-0">Pending</span>
                )}
                {result.status === 'checking' && (
                  <span className="text-[10px] text-amber-500 dark:text-amber-400 animate-pulse shrink-0">Checking…</span>
                )}
                {result.status === 'pass' && (
                  <span className="material-symbols-outlined text-green-500 text-base shrink-0">check_circle</span>
                )}
                {result.status === 'error' && (
                  <span className="material-symbols-outlined text-red-500 text-base shrink-0" title={result.message}>error</span>
                )}
                {result.status === 'issues' && (
                  <>
                    <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded shrink-0 ${
                      visibleMatches.length > 0
                        ? 'bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400'
                        : 'bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400'
                    }`}>
                      {visibleMatches.length > 0
                        ? `${visibleMatches.length} issue${visibleMatches.length > 1 ? 's' : ''}`
                        : 'Resolved ✓'}
                    </span>
                    <span className="material-symbols-outlined text-gray-400 text-sm shrink-0">
                      {isExpanded ? 'expand_less' : 'expand_more'}
                    </span>
                  </>
                )}
              </div>

              {/* Expanded issue details */}
              {result.status === 'issues' && isExpanded && (
                <ul className="border-t border-amber-100 dark:border-amber-900/30 bg-amber-50/40 dark:bg-amber-900/10 divide-y divide-amber-100 dark:divide-amber-900/20">
                  {result.matches.map((match, j) => {
                    const ignored = result.ignoredWords.has(match.word.toLowerCase())
                    return (
                      <li key={j} className={`flex items-start gap-2 px-4 py-1.5 ${ignored ? 'opacity-40' : ''}`}>
                        <div className="flex-1 min-w-0">
                          <span className="font-mono text-xs font-semibold text-red-600 dark:text-red-400">
                            &ldquo;{match.word}&rdquo;
                          </span>
                          <span className="text-xs text-gray-500 dark:text-gray-400 ml-1">— {match.message}</span>
                          {match.replacements.length > 0 && (
                            <span className="text-[10px] text-gray-400 dark:text-gray-500 ml-1">
                              → {match.replacements.slice(0, 3).join(', ')}
                            </span>
                          )}
                        </div>
                        {ignored ? (
                          <span className="text-[10px] text-gray-400 italic shrink-0">Ignored</span>
                        ) : (
                          <button
                            type="button"
                            onClick={e => { e.stopPropagation(); handleIgnoreWord(option.id, match.word) }}
                            className="shrink-0 text-[10px] px-1.5 py-0.5 text-gray-500 dark:text-gray-400 border border-gray-200 dark:border-gray-600 rounded hover:bg-gray-100 dark:hover:bg-gray-700 transition-colors"
                            title="Ignore this word in all future SPAG checks"
                          >
                            Ignore
                          </button>
                        )}
                      </li>
                    )
                  })}
                </ul>
              )}
            </li>
          )
        })}
      </ul>
    </div>
  )
}
