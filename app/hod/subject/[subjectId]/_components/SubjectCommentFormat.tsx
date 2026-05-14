"use client"

import { useState, useEffect } from "react"
import { updateSubjectCommentFormat } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"

const SAMPLE_DATA_BASE: Record<string, string> = {
  '<Name>': 'James',
  '<He>': 'He',
  '<he>': 'he',
  '<She>': 'He',
  '<she>': 'he',
  '<His>': 'His',
  '<his>': 'his',
  '<Her>': 'His',
  '<Him>': 'Him',
  '<him>': 'him',
  '<Year>': '7',
  '<EoYLevel>': 'B2',
  '<TargetLevel>': 'A1',
}

function pickRandom<T>(arr: T[]): T | undefined {
  if (arr.length === 0) return undefined
  return arr[Math.floor(Math.random() * arr.length)]
}

interface CommentOption {
  id: string
  code: string
  text: string
  displayOrder: number
}

interface CommentGroup {
  id: string
  name: string
  displayOrder: number
  CommentOption: CommentOption[]
}

interface Props {
  subjectId: string
  initialFormat: string | null
  groups: CommentGroup[]
  subjectTitle: string
}

export function SubjectCommentFormat({ subjectId, initialFormat, groups, subjectTitle }: Props) {
  const [format, setFormat] = useState(initialFormat || "")
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [seed, setSeed] = useState(0)
  const router = useRouter()

  async function handleSave() {
    setError(null)
    setSaving(true)
    setSaved(false)

    const result = await updateSubjectCommentFormat(subjectId, format || null)
    setSaving(false)

    if (!result.success) {
      setError(('error' in result ? result.error : "Failed to save") || "Failed to save")
    } else {
      setSaved(true)
      setTimeout(() => setSaved(false), 3000)
      router.refresh()
    }
  }

  const groupNames = groups.map(g => g.name)

  const [previewSegments, setPreviewSegments] = useState<{ text: string; unresolved: boolean }[]>([])

  // Build preview client-side only — pickRandom uses Math.random() which causes
  // SSR/client hydration mismatches if run during server render.
  useEffect(() => {
    if (!format) {
      setPreviewSegments([])
      return
    }

    const sampleData = { ...SAMPLE_DATA_BASE, '<Subject>': subjectTitle }

    const resolveVars = (text: string): string => {
      let r = text
      for (const [tag, value] of Object.entries(sampleData)) {
        r = r.replaceAll(tag, value)
      }
      return r
    }

    let result = format

    // Replace each <GroupName> tag with a random option's text (variables pre-resolved)
    for (const group of groups) {
      const option = pickRandom(group.CommentOption)
      const replacement = option ? resolveVars(option.text) : ''
      result = result.replaceAll(`<${group.name}>`, replacement)
    }

    // Resolve any remaining standard variables
    result = resolveVars(result)

    // Split into segments: resolved text and unresolved <Tag> portions
    const segments: { text: string; unresolved: boolean }[] = []
    const tagPattern = /<([^>]+)>/g
    let lastIndex = 0
    let match
    while ((match = tagPattern.exec(result)) !== null) {
      if (match.index > lastIndex) {
        segments.push({ text: result.slice(lastIndex, match.index), unresolved: false })
      }
      segments.push({ text: match[0], unresolved: true })
      lastIndex = match.index + match[0].length
    }
    if (lastIndex < result.length) {
      segments.push({ text: result.slice(lastIndex), unresolved: false })
    }

    setPreviewSegments(segments)
  }, [format, groups, seed, subjectTitle])

  return (
    <div className="bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm p-6">
      <div className="flex items-center gap-3 mb-4">
        <div className="size-8 rounded-lg bg-blue-100 dark:bg-blue-900/30 flex items-center justify-center text-blue-600">
          <span className="material-symbols-outlined text-lg">format_list_numbered</span>
        </div>
        <h2 className="text-lg font-bold text-[#111418] dark:text-white">Subject Comment Format</h2>
      </div>
      <p className="text-sm text-gray-500 mb-4">
        Define how this subject's comment groups are arranged using <code className="bg-gray-100 dark:bg-gray-800 px-1 rounded text-xs">&lt;GroupName&gt;</code> tags.
        Use separate lines for different paragraphs. Leave empty to join all groups in display order.
      </p>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Editor */}
        <div className="space-y-3">
          <textarea
            value={format}
            onChange={(e) => setFormat(e.target.value)}
            placeholder={"e.g.\n<WP> <TH>"}
            rows={5}
            className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-800 dark:text-white shadow-sm border p-3 text-sm font-mono"
          />
          {groupNames.length > 0 && (
            <div className="flex flex-wrap gap-1.5">
              <span className="text-xs text-gray-400">Available tags:</span>
              {groupNames.map(name => (
                <code key={name} className="text-xs bg-indigo-50 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-300 px-1.5 py-0.5 rounded font-mono">
                  &lt;{name}&gt;
                </code>
              ))}
            </div>
          )}
          <div className="flex items-center gap-3">
            <button
              onClick={handleSave}
              disabled={saving}
              className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 disabled:opacity-50 transition-colors"
            >
              {saving ? "Saving..." : "Save"}
            </button>
            {saved && <span className="text-sm text-green-600 font-medium">Saved!</span>}
            {error && <span className="text-sm text-red-600">{error}</span>}
          </div>
        </div>

        {/* Preview */}
        <div>
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-sm font-medium text-gray-700 dark:text-gray-300">Preview <span className="text-xs font-normal text-gray-400">(random options)</span></h3>
            <button
              onClick={() => setSeed(s => s + 1)}
              className="px-2 py-1 text-xs font-medium text-indigo-600 hover:text-indigo-800 dark:text-indigo-400 dark:hover:text-indigo-300 border border-indigo-200 dark:border-indigo-800 rounded hover:bg-indigo-50 dark:hover:bg-indigo-900/30 transition-colors"
            >
              Randomize
            </button>
          </div>
          <div className="bg-gray-50 dark:bg-gray-800 rounded-md p-4 text-sm whitespace-pre-wrap text-gray-700 dark:text-gray-300 min-h-[120px]">
            {previewSegments.length > 0 ? (
              previewSegments.map((seg, i) =>
                seg.unresolved ? (
                  <span key={i} className="text-red-500 font-medium">{seg.text}</span>
                ) : (
                  <span key={i}>{seg.text}</span>
                )
              )
            ) : (
              <span className="italic text-gray-400">Enter a format template to see a preview</span>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
