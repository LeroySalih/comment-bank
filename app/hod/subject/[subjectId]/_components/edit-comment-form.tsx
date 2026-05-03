'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { X } from 'lucide-react'
import { updateComment } from '@/lib/server-actions/hod'
import { requestSpagCheck, requestStandardsCheck } from '@/lib/server-actions/ai-check'
import { addIgnoredWord } from '@/lib/server-actions/ignored-words'
import { substituteVariables } from '@/lib/audit/substitute-variables'
import { VariablePreview } from '@/components/VariablePreview'
import { countWords } from '@/lib/utils'
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
  const [isSpagChecking, setIsSpagChecking] = useState(false)
  const [spagError, setSpagError] = useState<string | null>(null)
  // Track words ignored this session so they vanish from the list immediately
  const [locallyIgnored, setLocallyIgnored] = useState<Set<string>>(new Set())

  // Standards state
  const [standardsResult, setStandardsResult] = useState<StandardsResult | null>(null)
  const [isStandardsChecking, setIsStandardsChecking] = useState(false)
  const [standardsError, setStandardsError] = useState<string | null>(null)

  const router = useRouter()

  // Clear AI results when the comment text changes so stale results are not shown
  const handleTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(e.target.value)
    setSpagMatches(null)
    setStandardsResult(null)
    setSpagError(null)
    setStandardsError(null)
    setLocallyIgnored(new Set())
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
    setIsSpagChecking(true)
    setSpagMatches(null)
    setSpagError(null)
    setLocallyIgnored(new Set())

    // Substitute fixed test values so template variables like <Name> do not
    // trigger false SPAG positives.
    const substituted = substituteVariables(text, subjectTitle)
    const result = await requestSpagCheck(substituted)
    setIsSpagChecking(false)

    if (result.success) {
      setSpagMatches(result.result?.matches ?? [])
    } else {
      setSpagError(result.error)
    }
  }

  async function handleStandardsCheck() {
    if (!text.trim()) return
    setIsStandardsChecking(true)
    setStandardsResult(null)
    setStandardsError(null)

    const substituted = substituteVariables(text, subjectTitle)
    const result = await requestStandardsCheck(substituted)
    setIsStandardsChecking(false)

    if (result.success) {
      setStandardsResult(result.result)
    } else {
      setStandardsError(result.error)
    }
  }

  async function handleIgnoreWord(word: string) {
    // Persists to IgnoredWord table — automatically propagates to the audit
    // and pupil-report SPAG checks since they all read the same table.
    await addIgnoredWord(word)
    setLocallyIgnored(prev => new Set(prev).add(word.toLowerCase()))
  }

  // Filter out matches the teacher has already ignored this session
  const visibleMatches = (spagMatches ?? []).filter(
    m => !locallyIgnored.has(m.word.toLowerCase())
  )

  return (
    <div
      className="p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-200 dark:border-blue-800"
      onClick={e => e.stopPropagation()}
    >
      <div className="flex justify-between items-center mb-3">
        <h4 className="text-sm font-medium text-gray-900 dark:text-white">Edit Comment</h4>
        <button onClick={onClose} className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300">
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
          <div className="rounded-lg border border-amber-200 dark:border-amber-800 bg-white dark:bg-gray-900 overflow-hidden">
            <div className="flex items-center gap-2 px-3 py-2 bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-800">
              <span className="material-symbols-outlined text-amber-600 dark:text-amber-400 text-sm">
                spellcheck
              </span>
              <span className="text-xs font-semibold text-amber-700 dark:text-amber-400">
                {visibleMatches.length === 0
                  ? 'SPAG — No issues found ✓'
                  : `SPAG — ${visibleMatches.length} issue${visibleMatches.length > 1 ? 's' : ''} found`}
              </span>
            </div>
            {visibleMatches.length > 0 && (
              <ul className="divide-y divide-amber-100 dark:divide-amber-900/20">
                {visibleMatches.map(match => (
                  <li key={match.offset} className="px-3 py-2 flex items-start gap-2">
                    <div className="flex-1 min-w-0">
                      <span className="font-medium text-xs text-gray-800 dark:text-gray-200">
                        &ldquo;{match.word}&rdquo;
                      </span>
                      <span className="text-xs text-gray-500 dark:text-gray-400 ml-1">
                        — {match.message}
                      </span>
                      {match.replacements.length > 0 && (
                        <p className="text-[10px] text-gray-400 dark:text-gray-500 mt-0.5">
                          Suggestion{match.replacements.length > 1 ? 's' : ''}:{' '}
                          {match.replacements.slice(0, 3).join(', ')}
                        </p>
                      )}
                    </div>
                    <button
                      type="button"
                      onClick={() => handleIgnoreWord(match.word)}
                      className="flex-shrink-0 text-[10px] px-1.5 py-0.5 text-gray-500 dark:text-gray-400 hover:text-gray-700 dark:hover:text-gray-200 border border-gray-200 dark:border-gray-600 rounded hover:bg-gray-50 dark:hover:bg-gray-800 transition-colors"
                      title="Ignore this word in all future SPAG checks (audit and pupil reports)"
                    >
                      Ignore
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </div>
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
