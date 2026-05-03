'use client'

import { useState, useRef, useMemo, useEffect } from 'react'
import { useRouter } from 'next/navigation'
import { X } from 'lucide-react'
import { updateComment } from '@/lib/server-actions/hod'
import { requestSpagCheck } from '@/lib/server-actions/ai-check'
import { addIgnoredWord } from '@/lib/server-actions/ignored-words'
import { substituteVariables } from '@/lib/audit/substitute-variables'
import { VariablePreview } from '@/components/VariablePreview'
import { countWords } from '@/lib/utils'
import type { SpagResolution } from '@/components/SpagPanel'
import type { SpagMatch } from '@/lib/types/ai-check'

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

type Segment =
  | { type: 'unchanged'; text: string }
  | { type: 'match'; match: SpagMatch }

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

  // Two-stage save flow: 'clean' → no changes; 'dirty' → must check SPAG; 'spag_ok' → ready to save
  const [editStage, setEditStage] = useState<'clean' | 'dirty' | 'spag_ok'>('clean')

  // SPAG state
  const [spagMatches, setSpagMatches] = useState<SpagMatch[] | null>(null)
  const [spagResolutions, setSpagResolutions] = useState<Map<number, SpagResolution>>(new Map())
  const [spagOriginalText, setSpagOriginalText] = useState('')
  const [isSpagChecking, setIsSpagChecking] = useState(false)
  const [spagError, setSpagError] = useState<string | null>(null)

  // SPAG inline popup state
  const [activeSpagOffset, setActiveSpagOffset] = useState<number | null>(null)
  const [popupPlacement, setPopupPlacement] = useState<'above' | 'below'>('below')
  const [editingOffset, setEditingOffset] = useState<number | null>(null)
  const [editValue, setEditValue] = useState('')

  const spagRequestId = useRef(0)
  const spagContainerRef = useRef<HTMLDivElement | null>(null)
  const spagPopupRef = useRef<HTMLDivElement | null>(null)
  const spagEditInputRef = useRef<HTMLInputElement | null>(null)

  const router = useRouter()

  // Click outside closes the SPAG popup
  useEffect(() => {
    if (activeSpagOffset === null) return
    const handler = (e: MouseEvent) => {
      if (spagPopupRef.current && !spagPopupRef.current.contains(e.target as Node)) {
        const underline = spagContainerRef.current?.querySelector(
          `[data-spag-offset="${activeSpagOffset}"]`
        )
        if (underline && underline.contains(e.target as Node)) return
        setActiveSpagOffset(null)
        setEditingOffset(null)
        setEditValue('')
      }
    }
    document.addEventListener('mousedown', handler)
    return () => document.removeEventListener('mousedown', handler)
  }, [activeSpagOffset])

  // Focus edit input when edit mode opens
  useEffect(() => {
    if (editingOffset !== null && spagEditInputRef.current) {
      spagEditInputRef.current.focus()
      spagEditInputRef.current.select()
    }
  }, [editingOffset])

  // Segments: split spagOriginalText into unchanged / match slices
  const segments = useMemo<Segment[]>(() => {
    if (!spagMatches) return []
    const sorted = [...spagMatches].sort((a, b) => a.offset - b.offset)
    const out: Segment[] = []
    let cursor = 0
    for (const m of sorted) {
      if (m.offset > cursor) out.push({ type: 'unchanged', text: spagOriginalText.slice(cursor, m.offset) })
      out.push({ type: 'match', match: m })
      cursor = m.offset + m.length
    }
    if (cursor < spagOriginalText.length) out.push({ type: 'unchanged', text: spagOriginalText.slice(cursor) })
    return out
  }, [spagMatches, spagOriginalText])

  const unresolvedCount = spagMatches ? spagMatches.length - spagResolutions.size : 0

  // Clear SPAG results when the comment text changes; require re-check before saving
  const handleTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(e.target.value)
    spagRequestId.current++
    setSpagMatches(null)
    setSpagResolutions(new Map())
    setSpagError(null)
    setEditStage('dirty')
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

  // ── SPAG check ──────────────────────────────────────────────────────────────

  async function handleSpagCheck() {
    if (!text.trim()) return
    const id = ++spagRequestId.current
    setIsSpagChecking(true)
    setSpagMatches(null)
    setSpagResolutions(new Map())
    setSpagError(null)
    setActiveSpagOffset(null)

    const substituted = substituteVariables(text, subjectTitle)
    setSpagOriginalText(substituted)
    const result = await requestSpagCheck(substituted)

    if (spagRequestId.current !== id) return
    setIsSpagChecking(false)

    if (result.success) {
      const matches = result.result?.matches ?? []
      setSpagMatches(matches)
      // No issues at all → immediately ready to save
      if (matches.length === 0) setEditStage('spag_ok')
    } else {
      setSpagError(result.error)
    }
  }

  // Apply resolutions (reverse order so earlier offsets stay valid)
  function applyResolutions(original: string, resolutions: Map<number, SpagResolution>): string {
    const entries = [...resolutions.values()].sort((a, b) => b.match.offset - a.match.offset)
    let result = original
    for (const r of entries) {
      const start = r.match.offset
      const end = start + r.match.length
      const replacement =
        r.action === 'accepted' || r.action === 'edited' ? r.replacement
        : r.action === 'deleted' ? ''
        : r.match.word
      result = result.slice(0, start) + replacement + result.slice(end)
    }
    return result
  }

  function resolveSpag(resolution: SpagResolution, existingResolutions: Map<number, SpagResolution>) {
    if (resolution.action === 'ignored') {
      addIgnoredWord(resolution.match.word).catch(console.error)
    }
    const next = new Map(existingResolutions)
    next.set(resolution.match.offset, resolution)
    setSpagResolutions(next)
    if (resolution.action !== 'ignored') {
      setText(applyResolutions(spagOriginalText, next))
    }
    if (spagMatches && next.size === spagMatches.length) setEditStage('spag_ok')
    return next
  }

  function handleSpagAcceptAll() {
    if (!spagMatches) return
    const next = new Map(spagResolutions)
    for (const m of spagMatches) {
      if (next.has(m.offset)) continue
      const replacement = m.replacements[0]
      if (!replacement) continue
      next.set(m.offset, { match: m, action: 'accepted', replacement })
    }
    setSpagResolutions(next)
    setText(applyResolutions(spagOriginalText, next))
    if (next.size === spagMatches.length) setEditStage('spag_ok')
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
    if (next.size === spagMatches.length) setEditStage('spag_ok')
  }

  function clearSpag() {
    setSpagMatches(null)
    setSpagResolutions(new Map())
    setActiveSpagOffset(null)
  }

  function handleSpagRerun() {
    clearSpag()
    handleSpagCheck()
  }

  // ── SPAG popup handlers ─────────────────────────────────────────────────────

  function closePopup() {
    setActiveSpagOffset(null)
    setEditingOffset(null)
    setEditValue('')
  }

  function handleOpenPopup(e: React.MouseEvent<HTMLSpanElement>, match: SpagMatch) {
    const rect = e.currentTarget.getBoundingClientRect()
    const spaceBelow = window.innerHeight - rect.bottom
    setPopupPlacement(spaceBelow < 240 && rect.top > spaceBelow ? 'above' : 'below')
    setActiveSpagOffset(match.offset)
    setEditingOffset(null)
    setEditValue('')
  }

  function handleAccept(match: SpagMatch, replacement: string) {
    resolveSpag({ match, action: 'accepted', replacement }, spagResolutions)
    closePopup()
  }

  function handleEditStart(match: SpagMatch) {
    setEditingOffset(match.offset)
    setEditValue(match.replacements[0] ?? match.word)
  }

  function handleEditSubmit(match: SpagMatch) {
    resolveSpag({ match, action: 'edited', replacement: editValue }, spagResolutions)
    closePopup()
  }

  function handleDelete(match: SpagMatch) {
    resolveSpag({ match, action: 'deleted' }, spagResolutions)
    closePopup()
  }

  function handleIgnore(match: SpagMatch) {
    resolveSpag({ match, action: 'ignored' }, spagResolutions)
    closePopup()
  }

  function renderResolved(resolution: SpagResolution, key: string) {
    if (resolution.action === 'accepted' || resolution.action === 'edited') {
      return <span key={key} className="text-green-700 dark:text-green-400">{resolution.replacement}</span>
    }
    if (resolution.action === 'deleted') {
      return <span key={key} className="text-gray-400 dark:text-gray-500 italic">[deleted]</span>
    }
    return <span key={key} className="text-gray-600 dark:text-gray-300">{resolution.match.word}</span>
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

            {/* Textarea OR inline SPAG annotated view — same box */}
            {spagMatches !== null && !spagError ? (
              <div className="space-y-1">
                <div
                  ref={spagContainerRef}
                  className="block w-full rounded-md border border-amber-300 dark:border-amber-600 bg-white dark:bg-gray-700 shadow-sm p-2 text-sm leading-relaxed whitespace-pre-wrap min-h-[4.5rem]"
                >
                  {segments.map((seg, i) => {
                    if (seg.type === 'unchanged') return <span key={i}>{seg.text}</span>
                    const { match } = seg
                    const resolution = spagResolutions.get(match.offset)
                    if (resolution) return renderResolved(resolution, String(i))
                    const isActive = activeSpagOffset === match.offset
                    return (
                      <span key={i} className="relative inline-block">
                        <span
                          data-spag-offset={match.offset}
                          onClick={e => handleOpenPopup(e, match)}
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
                            ref={spagPopupRef}
                            className={`absolute z-50 left-0 ${popupPlacement === 'below' ? 'top-full mt-1' : 'bottom-full mb-1'} min-w-[14rem] max-w-[20rem] bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-700 rounded-lg shadow-lg p-2 text-sm font-sans whitespace-normal`}
                          >
                            <div className="flex items-start gap-2 px-1 pb-2 border-b border-gray-100 dark:border-gray-800">
                              <span className="material-symbols-outlined text-amber-500 text-base">spellcheck</span>
                              <div className="flex-1 min-w-0">
                                <p className="font-mono font-semibold text-red-600 dark:text-red-400 text-sm truncate">{match.word}</p>
                                <p className="text-xs text-gray-500 dark:text-gray-400 leading-snug">{match.message}</p>
                              </div>
                              <button type="button" onClick={closePopup} className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-200">
                                <span className="material-symbols-outlined text-sm">close</span>
                              </button>
                            </div>

                            {editingOffset === match.offset ? (
                              <div className="pt-2 px-1 flex items-center gap-1">
                                <input
                                  ref={spagEditInputRef}
                                  type="text"
                                  value={editValue}
                                  onChange={e => setEditValue(e.target.value)}
                                  onKeyDown={e => {
                                    if (e.key === 'Enter') { e.preventDefault(); handleEditSubmit(match) }
                                    else if (e.key === 'Escape') { e.preventDefault(); closePopup() }
                                  }}
                                  className="flex-1 px-2 py-1 border border-blue-400 rounded text-sm bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 focus:outline-none focus:ring-2 focus:ring-blue-300"
                                />
                                <button
                                  type="button"
                                  onClick={() => handleEditSubmit(match)}
                                  className="px-2 py-1 bg-green-600 hover:bg-green-700 text-white rounded text-xs font-semibold"
                                >
                                  Apply
                                </button>
                              </div>
                            ) : (
                              <>
                                {match.replacements.length > 0 ? (
                                  <div className="pt-2 px-1">
                                    <p className="text-[10px] uppercase font-semibold text-gray-400 mb-1">Suggestions</p>
                                    <div className="flex flex-wrap gap-1">
                                      {match.replacements.slice(0, 6).map(r => (
                                        <button
                                          key={r}
                                          type="button"
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
                                <div className="pt-2 px-1 mt-2 border-t border-gray-100 dark:border-gray-800 flex items-center gap-1">
                                  <button type="button" onClick={() => handleEditStart(match)} className="flex items-center gap-1 px-2 py-1 text-xs text-blue-600 dark:text-blue-400 hover:bg-blue-50 dark:hover:bg-blue-900/20 rounded">
                                    <span className="material-symbols-outlined text-sm">edit</span>Edit
                                  </button>
                                  <button type="button" onClick={() => handleDelete(match)} className="flex items-center gap-1 px-2 py-1 text-xs text-orange-600 dark:text-orange-400 hover:bg-orange-50 dark:hover:bg-orange-900/20 rounded">
                                    <span className="material-symbols-outlined text-sm">delete</span>Delete
                                  </button>
                                  <button type="button" onClick={() => handleIgnore(match)} className="flex items-center gap-1 px-2 py-1 text-xs text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-800 rounded ml-auto">
                                    <span className="material-symbols-outlined text-sm">visibility_off</span>Ignore
                                  </button>
                                </div>
                              </>
                            )}
                          </div>
                        )}
                      </span>
                    )
                  })}
                </div>

                {/* Compact SPAG action bar */}
                <div className="flex items-center gap-2">
                  <span className="text-[11px] text-amber-600 dark:text-amber-400">
                    {unresolvedCount > 0
                      ? `${unresolvedCount} issue${unresolvedCount > 1 ? 's' : ''} — click underlined word`
                      : 'All issues resolved ✓'}
                  </span>
                  {unresolvedCount > 0 && (
                    <>
                      <button type="button" onClick={handleSpagAcceptAll} className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium bg-green-600 hover:bg-green-700 text-white rounded transition-colors">
                        <span className="material-symbols-outlined text-xs">done_all</span>Accept all
                      </button>
                      <button type="button" onClick={handleSpagRejectAll} className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-medium border border-gray-300 dark:border-gray-600 text-gray-600 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-800 rounded transition-colors">
                        <span className="material-symbols-outlined text-xs">visibility_off</span>Ignore all
                      </button>
                    </>
                  )}
                  <button type="button" onClick={handleSpagRerun} className="flex items-center gap-1 px-2 py-0.5 text-[11px] text-amber-600 dark:text-amber-400 hover:bg-amber-50 dark:hover:bg-amber-900/20 rounded ml-auto transition-colors">
                    <span className="material-symbols-outlined text-xs">refresh</span>Re-run
                  </button>
                  <button type="button" onClick={clearSpag} className="flex items-center gap-1 px-2 py-0.5 text-[11px] text-gray-500 dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-gray-800 rounded transition-colors">
                    <span className="material-symbols-outlined text-xs">close</span>Dismiss
                  </button>
                </div>
              </div>
            ) : (
              <textarea
                value={text}
                onChange={handleTextChange}
                required
                rows={2}
                className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
              />
            )}
          </div>
        </div>

        {/* Male/Female preview */}
        <VariablePreview text={text} subjectName={subjectTitle} />

        {spagError && (
          <p className="text-xs text-red-600 dark:text-red-400">SPAG check failed: {spagError}</p>
        )}

        {/* Save / Cancel — primary button changes with edit stage */}
        <div className="flex gap-2 justify-end pt-1">
          <button
            type="button"
            onClick={onClose}
            className="bg-gray-200 dark:bg-gray-600 text-gray-800 dark:text-gray-200 px-3 py-1.5 rounded-md text-sm hover:bg-gray-300 dark:hover:bg-gray-500"
          >
            Cancel
          </button>

          {editStage === 'dirty' && spagMatches === null ? (
            /* Must check SPAG before saving */
            <button
              type="button"
              onClick={handleSpagCheck}
              disabled={isSpagChecking || !text.trim()}
              className="flex items-center gap-1.5 bg-amber-500 hover:bg-amber-600 disabled:opacity-50 disabled:cursor-not-allowed text-white px-3 py-1.5 rounded-md text-sm font-medium transition-colors"
            >
              <span className="material-symbols-outlined text-sm">spellcheck</span>
              {isSpagChecking ? 'Checking…' : 'Check SPAG'}
            </button>
          ) : (
            /* SPAG passed or issues all resolved → allow save */
            <button
              type="submit"
              disabled={saving || (editStage === 'dirty' && unresolvedCount > 0)}
              className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              {saving ? 'Saving…' : 'Save Changes'}
            </button>
          )}
        </div>

        {saveError && <p className="text-red-600 dark:text-red-400 text-xs">{saveError}</p>}
      </form>
    </div>
  )
}
