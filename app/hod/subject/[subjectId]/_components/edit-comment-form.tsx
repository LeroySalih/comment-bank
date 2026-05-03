'use client'

import { useState, useRef } from 'react'
import { useRouter } from 'next/navigation'
import { X } from 'lucide-react'
import { updateComment } from '@/lib/server-actions/hod'
import { requestSpagCheck, requestStandardsCheck } from '@/lib/server-actions/ai-check'
import { addIgnoredWord } from '@/lib/server-actions/ignored-words'
import { substituteVariables } from '@/lib/audit/substitute-variables'
import { VariablePreview } from '@/components/VariablePreview'
import { countWords } from '@/lib/utils'
import SpagPanel, { type SpagResolution } from '@/components/SpagPanel'
import type { SpagMatch, StandardsResult, StandardsRuleKey } from '@/lib/types/ai-check'

interface CommentOption {
  id: string
  code: string
  text: string
  displayOrder: number
}

interface EditCommentFormProps {
  comment: CommentOption
  subjectId: string
  groupId: string
  subjectTitle: string
  onClose: () => void
}

const STANDARDS_RULE_KEYS: StandardsRuleKey[] = [
  'UKSpelling',
  'CourseOverviewIncluded',
  'AcademicPerformanceIncluded',
  'TargetWordCountMet',
  'TerminologyCorrect',
  'JargonFree',
  'DataSpecific',
  'ToneBalanced',
  'SocialSkillsIncluded',
  'CollaborationIncluded',
  'BehaviourIncluded',
  'ParentalSupportIncluded',
  'Formatting',
]

export function EditCommentForm({
  comment,
  subjectId,
  groupId,
  subjectTitle,
  onClose,
}: EditCommentFormProps) {
  const [code, setCode] = useState(comment.code)
  const [text, setText] = useState(comment.text)
  const [saveError, setSaveError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)

  // SPAG state
  const [spagMatches, setSpagMatches] = useState<SpagMatch[] | null>(null)
  const [spagResolutions, setSpagResolutions] = useState<Map<number, SpagResolution>>(new Map())
  // The substituted text at the time the SPAG check ran — used as SpagPanel's source of truth
  const [spagOriginalText, setSpagOriginalText] = useState('')
  const [isSpagChecking, setIsSpagChecking] = useState(false)
  const [spagError, setSpagError] = useState<string | null>(null)

  // Standards state
  const [standardsResult, setStandardsResult] = useState<StandardsResult | null>(null)
  const [isStandardsChecking, setIsStandardsChecking] = useState(false)
  const [standardsError, setStandardsError] = useState<string | null>(null)

  const spagRequestId = useRef(0)
  const standardsRequestId = useRef(0)

  const router = useRouter()

  // Clear AI results when the comment text changes so stale results are not shown
  const handleTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(e.target.value)
    spagRequestId.current++
    standardsRequestId.current++
    setSpagMatches(null)
    setSpagResolutions(new Map())
    setStandardsResult(null)
    setSpagError(null)
    setStandardsError(null)
  }

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setSaveError(null)
    setSaving(true)

    const formData = new FormData()
    formData.append('code', code)
    formData.append('text', text)

    const result = await updateComment(comment.id, subjectId, groupId, formData)
    setSaving(false)

    if (!result.success) {
      setSaveError(
        ('error' in result ? result.error : 'Failed to update comment') ||
          'Failed to update comment'
      )
    } else {
      onClose()
      router.refresh()
    }
  }

  async function handleSpagCheck() {
    if (!text.trim()) return
    const id = ++spagRequestId.current
    setIsSpagChecking(true)
    setSpagMatches(null)
    setSpagResolutions(new Map())
    setSpagError(null)

    // Substitute fixed test values so template variables like <Name> do not
    // trigger false SPAG positives. Store the substituted text so SpagPanel
    // offsets remain valid even if the user edits the textarea afterwards.
    const substituted = substituteVariables(text, subjectTitle)
    setSpagOriginalText(substituted)
    const result = await requestSpagCheck(substituted)

    if (spagRequestId.current !== id) return // response is stale, discard it
    setIsSpagChecking(false)

    if (result.success) {
      setSpagMatches(result.result?.matches ?? [])
    } else {
      setSpagError(result.error)
    }
  }

  // Apply all resolutions (in reverse offset order) to produce updated text
  function applyResolutions(
    original: string,
    resolutions: Map<number, SpagResolution>
  ): string {
    const entries = [...resolutions.values()].sort((a, b) => b.match.offset - a.match.offset)
    let result = original
    for (const r of entries) {
      const start = r.match.offset
      const end = start + r.match.length
      const replacement =
        r.action === 'accepted' || r.action === 'edited'
          ? r.replacement
          : r.action === 'deleted'
          ? ''
          : r.match.word // ignored — keep word unchanged
      result = result.slice(0, start) + replacement + result.slice(end)
    }
    return result
  }

  function handleSpagResolve(resolution: SpagResolution) {
    if (resolution.action === 'ignored') {
      addIgnoredWord(resolution.match.word).catch(console.error)
    }
    const next = new Map(spagResolutions)
    next.set(resolution.match.offset, resolution)
    setSpagResolutions(next)

    // For text-changing actions, update the textarea so the teacher sees the fix
    if (resolution.action !== 'ignored') {
      setText(applyResolutions(spagOriginalText, next))
    }
  }

  function handleSpagAcceptAll() {
    if (!spagMatches) return
    const next = new Map(spagResolutions)
    for (const m of spagMatches) {
      if (next.has(m.offset)) continue
      const replacement = m.replacements[0]
      if (!replacement) continue // no suggestion — skip
      next.set(m.offset, { match: m, action: 'accepted', replacement })
    }
    setSpagResolutions(next)
    setText(applyResolutions(spagOriginalText, next))
  }

  function handleSpagRejectAll() {
    if (!spagMatches) return
    const next = new Map(spagResolutions)
    for (const m of spagMatches) {
      if (next.has(m.offset)) continue
      addIgnoredWord(m.word).catch(console.error)
      next.set(m.offset, { match: m, action: 'ignored' })
    }
    setSpagResolutions(next)
  }

  function handleSpagAllResolved() {
    setSpagMatches(null)
    setSpagResolutions(new Map())
  }

  function handleSpagDismiss() {
    setSpagMatches(null)
    setSpagResolutions(new Map())
  }

  function handleSpagRerun() {
    setSpagMatches(null)
    setSpagResolutions(new Map())
    handleSpagCheck()
  }

  async function handleStandardsCheck() {
    if (!text.trim()) return
    const id = ++standardsRequestId.current
    setIsStandardsChecking(true)
    setStandardsResult(null)
    setStandardsError(null)

    const substituted = substituteVariables(text, subjectTitle)
    const result = await requestStandardsCheck(substituted)

    if (standardsRequestId.current !== id) return // response is stale, discard it
    setIsStandardsChecking(false)

    if (result.success) {
      setStandardsResult(result.result)
    } else {
      setStandardsError(result.error)
    }
  }

  return (
    <div
      className="p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-200 dark:border-blue-800"
      onClick={e => e.stopPropagation()}
    >
      <div className="flex justify-between items-center mb-3">
        <h4 className="text-sm font-medium text-gray-900 dark:text-white">Edit Comment</h4>
        <button type="button" onClick={onClose} className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300">
          <X size={16} />
        </button>
      </div>

      <form onSubmit={handleSubmit} className="space-y-3">
        {/* Code + Text fields */}
        <div className="flex gap-3">
          <div className="w-24">
            <label className="block text-xs font-medium text-gray-500 dark:text-gray-400 mb-1">
              Code
            </label>
            <input
              type="text"
              value={code}
              onChange={e => setCode(e.target.value)}
              required
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
          <div className="flex-1">
            <div className="flex justify-between items-end mb-1">
              <label className="block text-xs font-medium text-gray-500 dark:text-gray-400">
                Comment Text
              </label>
              <span className="text-[10px] text-gray-400 font-medium">{countWords(text)} words</span>
            </div>
            <textarea
              value={text}
              onChange={handleTextChange}
              required
              rows={2}
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
        </div>

        {/* Male/Female preview — passes subjectTitle so <Subject> renders correctly */}
        <VariablePreview text={text} subjectName={subjectTitle} />

        {/* AI check buttons */}
        <div className="flex gap-2 flex-wrap">
          <button
            type="button"
            onClick={handleSpagCheck}
            disabled={isSpagChecking || !text.trim()}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-amber-700 dark:text-amber-400 bg-amber-50 dark:bg-amber-900/20 hover:bg-amber-100 dark:hover:bg-amber-900/40 border border-amber-200 dark:border-amber-800 rounded disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            <span className="material-symbols-outlined text-sm">spellcheck</span>
            {isSpagChecking ? 'Checking…' : 'Check SPAG'}
          </button>
          <button
            type="button"
            onClick={handleStandardsCheck}
            disabled={isStandardsChecking || !text.trim()}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-medium text-blue-700 dark:text-blue-400 bg-blue-50 dark:bg-blue-900/20 hover:bg-blue-100 dark:hover:bg-blue-900/40 border border-blue-200 dark:border-blue-800 rounded disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
          >
            <span className="material-symbols-outlined text-sm">checklist</span>
            {isStandardsChecking ? 'Checking…' : 'Check Standards'}
          </button>
        </div>

        {/* SPAG results */}
        {spagError && (
          <p className="text-xs text-red-600 dark:text-red-400">SPAG check failed: {spagError}</p>
        )}

        {spagMatches !== null && !spagError && (
          <SpagPanel
            originalText={spagOriginalText}
            matches={spagMatches}
            resolutions={spagResolutions}
            onResolve={handleSpagResolve}
            onAcceptAll={handleSpagAcceptAll}
            onRejectAll={handleSpagRejectAll}
            onAllResolved={handleSpagAllResolved}
            onRerun={handleSpagRerun}
            onDismiss={handleSpagDismiss}
          />
        )}

        {/* Standards results */}
        {standardsError && (
          <p className="text-xs text-red-600 dark:text-red-400">
            Standards check failed: {standardsError}
          </p>
        )}

        {standardsResult && !standardsError && (
          <div className="rounded-lg border border-blue-200 dark:border-blue-800 bg-white dark:bg-gray-900 overflow-hidden">
            <div className="flex items-center gap-2 px-3 py-2 bg-blue-50 dark:bg-blue-900/20 border-b border-blue-200 dark:border-blue-800">
              <span className="material-symbols-outlined text-blue-600 dark:text-blue-400 text-sm">
                checklist
              </span>
              <span className="text-xs font-semibold text-blue-700 dark:text-blue-400">
                Standards —{' '}
                {standardsResult.Status.result ? 'Passed ✓' : 'Issues found'}
              </span>
            </div>
            <ul className="divide-y divide-blue-100 dark:divide-blue-900/20">
              {STANDARDS_RULE_KEYS.map(key => {
                const entry = standardsResult[key]
                return (
                  <li
                    key={key}
                    className={`px-3 py-1.5 flex items-center gap-2 ${
                      !entry.result ? 'bg-red-50/50 dark:bg-red-900/10' : ''
                    }`}
                  >
                    <span
                      className={`material-symbols-outlined text-sm ${
                        entry.result ? 'text-green-500' : 'text-red-500'
                      }`}
                    >
                      {entry.result ? 'check_circle' : 'cancel'}
                    </span>
                    <span className="text-xs text-gray-700 dark:text-gray-300 flex-1">{key}</span>
                    {typeof entry.wordCount === 'number' && (
                      <span className="text-[10px] text-gray-400 dark:text-gray-500">
                        {entry.wordCount} words
                      </span>
                    )}
                  </li>
                )
              })}
            </ul>
          </div>
        )}

        {/* Save / Cancel */}
        <div className="flex gap-2 justify-end pt-1">
          <button
            type="button"
            onClick={onClose}
            className="bg-gray-200 dark:bg-gray-600 text-gray-800 dark:text-gray-200 px-3 py-1.5 rounded-md text-sm hover:bg-gray-300 dark:hover:bg-gray-500"
          >
            Cancel
          </button>
          <button
            type="submit"
            disabled={saving}
            className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700 disabled:opacity-50"
          >
            {saving ? 'Saving…' : 'Save Changes'}
          </button>
        </div>

        {saveError && <p className="text-red-600 dark:text-red-400 text-xs">{saveError}</p>}
      </form>
    </div>
  )
}
