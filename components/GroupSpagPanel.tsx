'use client'

import { useState, useEffect, useRef, useMemo } from 'react'
import { createPortal } from 'react-dom'
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
  | { status: 'issues'; matches: SpagMatch[]; ignoredWords: Set<string>; spagText: string }
  | { status: 'error'; message: string }

interface GroupSpagPanelProps {
  options: Option[]
  subjectTitle?: string
  onClose: () => void
}

// ── Inline annotation renderer (read-only) ────────────────────────────────────

type Segment =
  | { type: 'unchanged'; text: string }
  | { type: 'match'; match: SpagMatch }

interface PopupState {
  match: SpagMatch
  top?: number
  bottom?: number
  left: number
}

function SpagAnnotatedText({
  spagText,
  matches,
  ignoredWords,
  onIgnore,
}: {
  spagText: string
  matches: SpagMatch[]
  ignoredWords: Set<string>
  onIgnore: (word: string) => void
}) {
  const [popup, setPopup] = useState<PopupState | null>(null)
  const [mounted, setMounted] = useState(false)
  const popupRef = useRef<HTMLDivElement>(null)

  useEffect(() => { setMounted(true) }, [])

  // Close popup on outside click
  useEffect(() => {
    if (!popup) return
    const handler = (e: MouseEvent) => {
      if (popupRef.current && !popupRef.current.contains(e.target as Node)) {
        setPopup(null)
      }
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [popup])

  const visibleMatches = useMemo(
    () => matches.filter(m => !ignoredWords.has(m.word.toLowerCase())),
    [matches, ignoredWords]
  )

  const segments = useMemo<Segment[]>(() => {
    const sorted = [...visibleMatches].sort((a, b) => a.offset - b.offset)
    const out: Segment[] = []
    let cursor = 0
    for (const m of sorted) {
      if (m.offset > cursor) out.push({ type: 'unchanged', text: spagText.slice(cursor, m.offset) })
      out.push({ type: 'match', match: m })
      cursor = m.offset + m.length
    }
    if (cursor < spagText.length) out.push({ type: 'unchanged', text: spagText.slice(cursor) })
    return out
  }, [visibleMatches, spagText])

  function handleWordClick(e: React.MouseEvent<HTMLSpanElement>, match: SpagMatch) {
    e.stopPropagation()
    // Toggle off if same word
    if (popup?.match.offset === match.offset) {
      setPopup(null)
      return
    }
    const rect = e.currentTarget.getBoundingClientRect()
    const spaceBelow = window.innerHeight - rect.bottom
    const newPopup: PopupState = { match, left: Math.min(rect.left, window.innerWidth - 320) }
    if (spaceBelow >= 180 || spaceBelow > rect.top) {
      newPopup.top = rect.bottom + 4
    } else {
      newPopup.bottom = window.innerHeight - rect.top + 4
    }
    setPopup(newPopup)
  }

  // Portal popup — renders at document.body level, bypasses all overflow:hidden
  const portalContent = mounted && popup ? createPortal(
    <div
      ref={popupRef}
      style={{
        position: 'fixed',
        top: popup.top,
        bottom: popup.bottom,
        left: popup.left,
        zIndex: 9999,
      }}
      className="min-w-[14rem] max-w-[20rem] bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-lg shadow-xl p-2 text-sm font-sans whitespace-normal"
    >
      {/* Header */}
      <div className="flex items-start gap-2 px-1 pb-2 border-b border-gray-100 dark:border-gray-800">
        <span className="material-symbols-outlined text-amber-500 text-base">spellcheck</span>
        <div className="flex-1 min-w-0">
          <p className="font-mono font-semibold text-red-600 dark:text-red-400 text-sm truncate">
            &ldquo;{popup.match.word}&rdquo;
          </p>
          <p className="text-xs text-gray-500 dark:text-gray-400 leading-snug">{popup.match.message}</p>
        </div>
        <button
          type="button"
          onClick={() => setPopup(null)}
          className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-200"
        >
          <span className="material-symbols-outlined text-sm">close</span>
        </button>
      </div>

      {/* Suggestions (read-only display) */}
      {popup.match.replacements.length > 0 ? (
        <div className="pt-2 px-1">
          <p className="text-[10px] uppercase font-semibold text-gray-400 mb-1">Suggestions</p>
          <div className="flex flex-wrap gap-1">
            {popup.match.replacements.slice(0, 6).map(r => (
              <span
                key={r}
                className="px-2 py-1 bg-gray-50 dark:bg-gray-800 border border-gray-200 dark:border-gray-600 rounded text-xs font-mono text-gray-700 dark:text-gray-300"
              >
                {r}
              </span>
            ))}
          </div>
        </div>
      ) : (
        <p className="pt-2 px-1 text-xs text-gray-400 italic">No suggestions available</p>
      )}

      {/* Ignore action */}
      <div className="pt-2 px-1 mt-2 border-t border-gray-100 dark:border-gray-800 flex">
        <button
          type="button"
          onClick={() => { onIgnore(popup.match.word); setPopup(null) }}
          className="flex items-center gap-1 px-2 py-1 text-xs text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-800 rounded ml-auto"
          title="Ignore this word in all future SPAG checks"
        >
          <span className="material-symbols-outlined text-sm">visibility_off</span>Ignore
        </button>
      </div>
    </div>,
    document.body
  ) : null

  return (
    <>
      <div className="text-xs text-gray-700 dark:text-gray-300 leading-relaxed">
        {segments.map((seg, i) => {
          if (seg.type === 'unchanged') return <span key={i}>{seg.text}</span>
          const { match } = seg
          const isActive = popup?.match.offset === match.offset
          return (
            <span
              key={i}
              data-spag-offset={match.offset}
              onClick={e => handleWordClick(e, match)}
              className={`cursor-pointer ${isActive ? 'text-red-800 dark:text-red-300' : 'text-red-700 dark:text-red-400'}`}
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
          )
        })}
      </div>
      {portalContent}
    </>
  )
}

// ── Main panel ────────────────────────────────────────────────────────────────

export function GroupSpagPanel({ options, subjectTitle, onClose }: GroupSpagPanelProps) {
  const [results, setResults] = useState<Map<string, ItemResult>>(
    () => new Map(options.map(o => [o.id, { status: 'pending' }]))
  )
  const [isRunning, setIsRunning] = useState(false)

  // Start checking automatically when the panel mounts
  useEffect(() => { runChecks() }, []) // eslint-disable-line react-hooks/exhaustive-deps

  async function runChecks() {
    setIsRunning(true)
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
            : { status: 'issues', matches, ignoredWords: new Set(), spagText: substituted }
        ))
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

  const total = options.length
  const done = [...results.values()].filter(r => r.status !== 'pending' && r.status !== 'checking').length
  const issueItems = [...results.values()].filter(r => r.status === 'issues')
  const totalIssues = issueItems.reduce((sum, r) =>
    r.status === 'issues' ? sum + r.matches.filter(m => !r.ignoredWords.has(m.word.toLowerCase())).length : sum
  , 0)

  return (
    <div className="mb-4 border border-amber-200 dark:border-amber-800 rounded-lg bg-white dark:bg-gray-900">
      {/* Header */}
      <div className="flex items-center gap-2 px-3 py-2 bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-800 rounded-t-lg">
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
              onClick={() => runChecks()}
              className="flex items-center gap-1 text-xs text-amber-600 dark:text-amber-400 hover:text-amber-700 dark:hover:text-amber-300 transition-colors"
            >
              <span className="material-symbols-outlined text-sm">refresh</span>Re-run
            </button>
          )}
          <button
            type="button"
            onClick={onClose}
            className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-200"
          >
            <span className="material-symbols-outlined text-sm">close</span>
          </button>
        </div>
      </div>

      {/* Per-item results */}
      <ul className="divide-y divide-gray-100 dark:divide-gray-800">
        {options.map(option => {
          const result = results.get(option.id) ?? { status: 'pending' as const }
          const visibleMatchCount = result.status === 'issues'
            ? result.matches.filter(m => !result.ignoredWords.has(m.word.toLowerCase())).length
            : 0

          return (
            <li key={option.id} className="px-3 py-2.5">
              {/* Code + status badge */}
              <div className="flex items-center gap-2 mb-1">
                <span className="text-xs font-mono font-bold text-gray-500 dark:text-gray-400 w-8 shrink-0">
                  {option.code}
                </span>

                {result.status === 'pending' && (
                  <span className="text-[10px] text-gray-400">Pending</span>
                )}
                {result.status === 'checking' && (
                  <span className="text-[10px] text-amber-500 dark:text-amber-400 animate-pulse">Checking…</span>
                )}
                {result.status === 'pass' && (
                  <span className="material-symbols-outlined text-green-500 text-base leading-none">check_circle</span>
                )}
                {result.status === 'error' && (
                  <span className="material-symbols-outlined text-red-500 text-base leading-none" title={result.message}>
                    error
                  </span>
                )}
                {result.status === 'issues' && (
                  <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded ${
                    visibleMatchCount > 0
                      ? 'bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400'
                      : 'bg-green-100 dark:bg-green-900/30 text-green-600 dark:text-green-400'
                  }`}>
                    {visibleMatchCount > 0
                      ? `${visibleMatchCount} issue${visibleMatchCount > 1 ? 's' : ''}`
                      : 'Resolved ✓'}
                  </span>
                )}
              </div>

              {/* Text — annotated inline if issues, plain otherwise */}
              <div className="pl-10">
                {result.status === 'issues' ? (
                  <SpagAnnotatedText
                    spagText={result.spagText}
                    matches={result.matches}
                    ignoredWords={result.ignoredWords}
                    onIgnore={word => handleIgnoreWord(option.id, word)}
                  />
                ) : (
                  <p className="text-xs text-gray-600 dark:text-gray-400">
                    {option.text.length > 120 ? option.text.slice(0, 120) + '…' : option.text}
                  </p>
                )}
              </div>
            </li>
          )
        })}
      </ul>
    </div>
  )
}
