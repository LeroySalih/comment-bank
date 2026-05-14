'use client'

import { useState, useEffect, useRef, useMemo } from 'react'
import { requestSpagCheck, requestStandardsCheck } from '@/lib/server-actions/ai-check'
import { addIgnoredWord } from '@/lib/server-actions/ignored-words'
import { substituteVariables } from '@/lib/audit/substitute-variables'
import type { StandardsResult, StandardsRuleKey } from '@/lib/types/ai-check'

const DEBOUNCE_MS = 3000

// ─── SPAG ──────────────────────────────────────────────────────────────────────

type IssueType = 'spelling' | 'grammar' | 'style'

interface Issue {
  offset: number
  length: number
  word: string
  message: string
  replacements: string[]
  type: IssueType
}

const ISSUE_STYLE: Record<IssueType, { underline: string; bg: string }> = {
  spelling: { underline: '#ef4444', bg: '#fef2f2' },
  grammar:  { underline: '#3b82f6', bg: '#eff6ff' },
  style:    { underline: '#8b5cf6', bg: '#f5f3ff' },
}

function classifyType(message: string): IssueType {
  const m = message.toLowerCase()
  if (m.includes('spell') || m.includes('typo') || m.includes('misspell')) return 'spelling'
  if (m.includes('style') || m.includes('wordy') || m.includes('passive')) return 'style'
  return 'grammar'
}

// ─── Standards ─────────────────────────────────────────────────────────────────

const STANDARDS_RULES: { key: StandardsRuleKey; label: string }[] = [
  { key: 'UKSpelling',                label: 'UK Spelling' },
  { key: 'CourseOverviewIncluded',    label: 'Course Overview' },
  { key: 'TargetWordCountMet',        label: 'Word Count' },
  { key: 'TerminologyCorrect',        label: 'Terminology' },
  { key: 'JargonFree',                label: 'Jargon Free' },
  { key: 'DataSpecific',              label: 'Data Specific' },
  { key: 'ToneBalanced',              label: 'Tone Balanced' },
  { key: 'SocialSkillsIncluded',      label: 'Social Skills' },
  { key: 'CollaborationIncluded',     label: 'Collaboration' },
  { key: 'BehaviourIncluded',         label: 'Behaviour' },
  { key: 'ParentalSupportIncluded',   label: 'Parental Support' },
  { key: 'Formatting',                label: 'Formatting' },
]

// ─── Shared font — textarea and overlay must match exactly ─────────────────────

const SHARED: React.CSSProperties = {
  fontFamily: 'ui-sans-serif, system-ui, -apple-system, "Segoe UI", sans-serif',
  fontSize: '14px',
  lineHeight: '1.5',
  letterSpacing: 'normal',
  padding: '8px 10px',
  wordBreak: 'break-word',
  whiteSpace: 'pre-wrap',
  overflowWrap: 'break-word',
}

// ─── Component ─────────────────────────────────────────────────────────────────

export interface InlineSpagEditorProps {
  value: string
  onChange: (value: string) => void
  subjectTitle?: string
  rows?: number
  placeholder?: string
  name?: string
  required?: boolean
  spagOnly?: boolean
}

export function InlineSpagEditor({
  value,
  onChange,
  subjectTitle = '',
  rows = 3,
  placeholder,
  name,
  required,
  spagOnly = false,
}: InlineSpagEditorProps) {
  // SPAG state
  const [issues, setIssues] = useState<Issue[]>([])
  const [spagChecking, setSpagChecking] = useState(false)
  const [dismissed, setDismissed] = useState<Set<number>>(new Set())
  const spagReqRef = useRef(0)
  // rawValueAtCheck: the value prop at the moment the last SPAG result arrived
  // checkedValue: the substituted string sent to LanguageTool (offsets refer to this)
  const rawValueAtCheckRef = useRef('')
  const checkedValueRef = useRef('')
  // Set true before applying a suggestion so the resulting onChange doesn't re-trigger the check
  const skipNextSpagRef = useRef(false)

  // Standards state
  const [standardsResult, setStandardsResult] = useState<StandardsResult | null>(null)
  const [standardsRaw, setStandardsRaw] = useState<unknown>(null)
  const [standardsChecking, setStandardsChecking] = useState(false)
  const stdReqRef = useRef(0)

  // Auto SPAG: 3s after typing stops (skipped when value changed via suggestion apply)
  useEffect(() => {
    if (skipNextSpagRef.current) { skipNextSpagRef.current = false; return }
    if (!value.trim()) {
      setIssues([])
      setSpagChecking(false)
      return
    }
    setSpagChecking(true)
    const myReq = ++spagReqRef.current
    const timer = setTimeout(async () => {
      const toCheck = subjectTitle ? substituteVariables(value, subjectTitle) : value
      const res = await requestSpagCheck(toCheck)
      if (myReq !== spagReqRef.current) return
      // Record which value these results belong to before updating state
      checkedValueRef.current = toCheck
      rawValueAtCheckRef.current = value
      setSpagChecking(false)
      if (res.success && res.result) {
        setIssues(
          res.result.matches.map((m) => ({
            offset: m.offset,
            length: m.length,
            word: m.word,
            message: m.message,
            replacements: m.replacements,
            type: classifyType(m.message),
          }))
        )
        setDismissed(new Set())
      } else {
        setIssues([])
      }
    }, DEBOUNCE_MS)
    return () => clearTimeout(timer)
  }, [value, subjectTitle])

  // Auto Standards: 3s after typing stops (skipped when spagOnly)
  useEffect(() => {
    if (spagOnly) return
    if (!value.trim()) {
      setStandardsResult(null)
      setStandardsChecking(false)
      return
    }
    setStandardsChecking(true)
    const myReq = ++stdReqRef.current
    const timer = setTimeout(async () => {
      const res = await requestStandardsCheck(value)
      if (myReq !== stdReqRef.current) return
      setStandardsChecking(false)
      if (res.success) {
        setStandardsResult(res.result ?? null)
        setStandardsRaw(res.raw)
      } else {
        setStandardsResult(null)
        setStandardsRaw(null)
      }
    }, DEBOUNCE_MS)
    return () => clearTimeout(timer)
  }, [value, spagOnly])

  // Issues are only valid when value matches what was checked — discard if user has typed since
  const activeIssues = useMemo(
    () => value === rawValueAtCheckRef.current
      ? issues.filter((iss) => !dismissed.has(iss.offset))
      : [],
    [value, issues, dismissed]
  )

  const failingRules = useMemo(() => {
    if (!standardsResult) return []
    return STANDARDS_RULES.filter((r) => standardsResult[r.key]?.result === false)
  }, [standardsResult])

  // Build overlay segments
  const segments = useMemo(() => {
    const src = checkedValueRef.current || value
    const sorted = [...activeIssues].sort((a, b) => a.offset - b.offset)
    const segs: Array<{ text: string; issue?: Issue }> = []
    let cursor = 0
    for (const iss of sorted) {
      if (iss.offset > cursor) segs.push({ text: src.slice(cursor, iss.offset) })
      segs.push({ text: src.slice(iss.offset, iss.offset + iss.length), issue: iss })
      cursor = iss.offset + iss.length
    }
    if (cursor < src.length) segs.push({ text: src.slice(cursor) })
    if (src.endsWith('\n')) segs.push({ text: '​' })
    return segs
  }, [value, activeIssues])

  function applyReplacement(iss: Issue, replacement: string) {
    const src = checkedValueRef.current || value
    const next = src.slice(0, iss.offset) + replacement + src.slice(iss.offset + iss.length)
    // Keep refs in sync with the incoming value so the stale-guard keeps remaining issues visible
    rawValueAtCheckRef.current = next
    checkedValueRef.current = next
    skipNextSpagRef.current = true
    onChange(next)
    setIssues((prev) => prev.filter((i) => i.offset !== iss.offset))
  }

  function dismissIssue(iss: Issue) {
    setDismissed((prev) => new Set([...prev, iss.offset]))
  }

  function deleteWord(iss: Issue) {
    const src = checkedValueRef.current || value
    const next = src.slice(0, iss.offset) + src.slice(iss.offset + iss.length)
    rawValueAtCheckRef.current = next
    checkedValueRef.current = next
    skipNextSpagRef.current = true
    onChange(next)
    setIssues((prev) => prev.filter((i) => i.offset !== iss.offset))
  }

  function ignoreWord(iss: Issue) {
    addIgnoredWord(iss.word).catch(console.error)
    const lower = iss.word.toLowerCase()
    setIssues((prev) => prev.filter((i) => i.word.toLowerCase() !== lower))
  }

  const checking = spagChecking || standardsChecking

  return (
    <div>
      {/* Editor box */}
      <div className="relative rounded-md border border-gray-300 dark:border-gray-600 bg-white dark:bg-gray-700 focus-within:ring-2 focus-within:ring-indigo-400 focus-within:border-transparent overflow-hidden">
        {/* Highlight overlay */}
        <div
          aria-hidden
          style={{
            ...SHARED,
            position: 'absolute',
            inset: 0,
            pointerEvents: 'none',
            color: 'transparent',
            overflow: 'hidden',
          }}
        >
          {segments.map((seg, i) =>
            seg.issue ? (
              <span
                key={i}
                style={{
                  textDecoration: `underline wavy ${ISSUE_STYLE[seg.issue.type].underline}`,
                  textDecorationThickness: '1.5px',
                  textUnderlineOffset: '3px',
                }}
              >
                {seg.text}
              </span>
            ) : (
              <span key={i}>{seg.text}</span>
            )
          )}
        </div>

        {/* Textarea */}
        <textarea
          name={name}
          required={required}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          rows={rows}
          placeholder={placeholder}
          style={{
            ...SHARED,
            position: 'relative',
            display: 'block',
            width: '100%',
            background: 'transparent',
            color: 'inherit',
            caretColor: 'currentColor',
            resize: 'vertical',
            outline: 'none',
            border: 'none',
          }}
          className="dark:text-white placeholder:text-gray-400"
        />
      </div>

      {/* Checking spinner */}
      {checking && activeIssues.length === 0 && failingRules.length === 0 && (
        <div className="mt-1 flex items-center gap-1.5 text-[11px] text-gray-400 select-none">
          <span
            className="inline-block w-2.5 h-2.5 rounded-full border-2 border-gray-300 border-t-gray-500 animate-spin"
            style={{ animationDuration: '0.7s' }}
          />
          Checking spelling, grammar &amp; standards…
        </div>
      )}

      {/* SPAG issue rows */}
      {activeIssues.length > 0 && (
        <div className="mt-1.5 flex flex-col gap-1">
          {activeIssues.map((iss) => (
            <IssueRow
              key={`${iss.offset}-${iss.word}`}
              issue={iss}
              onApply={(r) => applyReplacement(iss, r)}
              onDelete={() => deleteWord(iss)}
              onDismiss={() => dismissIssue(iss)}
              onIgnore={() => ignoreWord(iss)}
            />
          ))}
        </div>
      )}

      {/* Standards failing rules */}
      {failingRules.length > 0 && (
        <div className="mt-2">
          <div className="flex items-center gap-1.5 mb-1">
            <span className="text-[10px] font-semibold uppercase tracking-wide text-amber-600 dark:text-amber-400">
              Standards
            </span>
            {standardsChecking && (
              <span
                className="inline-block w-2 h-2 rounded-full border-2 border-amber-200 border-t-amber-500 animate-spin"
                style={{ animationDuration: '0.7s' }}
              />
            )}
          </div>
          <div className="flex flex-col gap-2">
            <div className="flex flex-wrap gap-1.5">
              {failingRules.map((r) => (
                <span
                  key={r.key}
                  className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium bg-amber-50 dark:bg-amber-900/30 text-amber-700 dark:text-amber-400 border border-amber-200 dark:border-amber-700"
                >
                  <svg width="9" height="9" viewBox="0 0 9 9" fill="currentColor">
                    <circle cx="4.5" cy="4.5" r="4.5" opacity="0.2" />
                    <circle cx="4.5" cy="4.5" r="2" />
                  </svg>
                  {r.label}
                </span>
              ))}
            </div>
            {/* Feedback details for failing rules */}
            {failingRules.map((r) => {
              const entry = standardsResult?.[r.key]
              if (!entry) return null
              const hasWordCount = typeof entry.wordCount === 'number'
              const hasInstances = entry.instances && entry.instances.length > 0
              if (!hasWordCount && !hasInstances) return null
              return (
                <div key={r.key} className="rounded border border-amber-200 dark:border-amber-800 bg-amber-50 dark:bg-amber-900/20 px-2.5 py-2 text-[11px] text-amber-800 dark:text-amber-300">
                  <p className="font-semibold mb-1">{r.label}</p>
                  {hasWordCount && (
                    <p>{entry.wordCount} words — target 200–300</p>
                  )}
                  {hasInstances && (
                    <ul className="mt-1 space-y-0.5 list-disc list-inside">
                      {entry.instances!.map((inst, i) => (
                        <li key={i} className="font-mono truncate">{inst}</li>
                      ))}
                    </ul>
                  )}
                </div>
              )
            })}
          </div>
        </div>
      )}

      {/* Debug: raw standards service response */}
      {standardsRaw !== null && (
        <details className="mt-2">
          <summary className="text-[10px] text-gray-400 cursor-pointer select-none">Standards service response (raw)</summary>
          <pre className="mt-1 text-[10px] bg-gray-50 dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded p-2 overflow-auto max-h-64 text-gray-600 dark:text-gray-400">
            {JSON.stringify(standardsRaw, null, 2)}
          </pre>
        </details>
      )}

      {/* All clear */}
      {!checking && standardsResult && failingRules.length === 0 && activeIssues.length === 0 && (
        <div className="mt-1 text-[11px] text-green-600 dark:text-green-400 flex items-center gap-1">
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="2 6 5 9 10 3" />
          </svg>
          All checks passed
        </div>
      )}
    </div>
  )
}

// ─── IssueRow ──────────────────────────────────────────────────────────────────

function IssueRow({
  issue,
  onApply,
  onDelete,
  onDismiss,
  onIgnore,
}: {
  issue: Issue
  onApply: (r: string) => void
  onDelete: () => void
  onDismiss: () => void
  onIgnore: () => void
}) {
  const s = ISSUE_STYLE[issue.type]
  return (
    <div
      className="rounded px-2 py-1.5 flex items-center gap-2 text-xs dark:bg-gray-800"
      style={{ background: s.bg }}
    >
      <span
        className="shrink-0 font-mono font-semibold"
        style={{
          color: s.underline,
          textDecoration: `underline wavy ${s.underline}`,
          textDecorationThickness: '1.5px',
          textUnderlineOffset: '2px',
        }}
      >
        {issue.word}
      </span>

      <span className="flex-1 min-w-0 truncate text-gray-500 dark:text-gray-400">
        {issue.message}
      </span>

      <div className="flex items-center gap-1 shrink-0 flex-wrap">
        {issue.replacements.slice(0, 3).map((r) => (
          <button
            key={r}
            type="button"
            onClick={() => onApply(r)}
            className="px-1.5 py-0.5 bg-white dark:bg-gray-700 border border-gray-200 dark:border-gray-600 rounded text-[11px] font-mono text-gray-800 dark:text-gray-200 hover:bg-green-50 dark:hover:bg-green-900/30 hover:border-green-300 transition-colors"
          >
            {r}
          </button>
        ))}

        <button
          type="button"
          onClick={onDelete}
          title="Delete word"
          className="ml-0.5 text-gray-400 hover:text-red-500 dark:hover:text-red-400 leading-none"
        >
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="3 6 5 6 21 6" /><path d="M19 6l-1 14H6L5 6" /><path d="M10 11v6M14 11v6" /><path d="M9 6V4h6v2" />
          </svg>
        </button>

        <button
          type="button"
          onClick={onDismiss}
          title="Dismiss"
          className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 leading-none"
        >
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round">
            <line x1="2" y1="2" x2="10" y2="10" /><line x1="10" y1="2" x2="2" y2="10" />
          </svg>
        </button>

        <button
          type="button"
          onClick={onIgnore}
          title="Ignore this word everywhere"
          className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 leading-none"
        >
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
            <path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94" />
            <path d="M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19" />
            <line x1="1" y1="1" x2="23" y2="23" />
          </svg>
        </button>
      </div>
    </div>
  )
}
