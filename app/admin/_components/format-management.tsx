"use client"

import { useState, useMemo } from "react"
import { updateCommentFormatTemplate } from "@/lib/server-actions/ccg"

const SAMPLE_DATA: Record<string, string> = {
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

interface SubjectCommentGroup {
  id: string
  name: string
  displayOrder: number
  CommentOption: { id: string; code: string; text: string }[]
}

interface SubjectWithGroups {
  id: string
  code: string
  title?: string | null
  commentFormat?: string | null
  CommentGroup?: SubjectCommentGroup[]
}

interface Props {
  initialTemplate: string
  commonGroups: { id: string; name: string; CommonCommentOption: { id: string; code: string; text: string; displayOrder: number }[] }[]
  subjects: SubjectWithGroups[]
}

export function FormatManagement({ initialTemplate, commonGroups, subjects }: Props) {
  const [template, setTemplate] = useState(initialTemplate)
  const [saving, setSaving] = useState(false)
  const [saved, setSaved] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [previewSubjectId, setPreviewSubjectId] = useState<string>(subjects[0]?.id || '')
  const [seed, setSeed] = useState(0) // bump to re-randomize

  async function handleSave() {
    setError(null)
    setSaving(true)
    setSaved(false)

    const result = await updateCommentFormatTemplate(template)
    setSaving(false)

    if (!result.success) {
      setError(('error' in result ? result.error : "Failed to save template") || "Failed to save template")
    } else {
      setSaved(true)
      setTimeout(() => setSaved(false), 3000)
    }
  }

  const availableTags = commonGroups.map(g => `<${g.name}>`)

  const selectedSubject = subjects.find(s => s.id === previewSubjectId)

  // Build preview segments: { text, unresolved } pairs
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const previewSegments = useMemo(() => {
    if (!template) return []

    const subjectTitle = selectedSubject
      ? (selectedSubject.title || selectedSubject.code)
      : ''

    // Resolve standard variables within an option's text before inserting into template
    const resolveVars = (text: string): string => {
      let r = text
      r = r.replaceAll('<Subject>', subjectTitle)
      for (const [tag, value] of Object.entries(SAMPLE_DATA)) {
        r = r.replaceAll(tag, value)
      }
      return r
    }

    let result = template

    // Replace each CCG group tag with a random option's text (variables pre-resolved)
    for (const group of commonGroups) {
      const options = group.CommonCommentOption || []
      const option = pickRandom(options)
      const replacement = option ? resolveVars(option.text) : ''
      result = result.replaceAll(`<${group.name}>`, replacement)
    }

    // Replace <SCG> tag using the selected subject's comment groups
    if (selectedSubject) {
      const groups = selectedSubject.CommentGroup || []
      let subjectParts: string[]
      if (selectedSubject.commentFormat && groups.length > 0) {
        const codes = selectedSubject.commentFormat.split(/\s+/)
        subjectParts = codes.map(code => {
          const g = groups.find(g => g.name === code)
          if (!g) return ''
          const options = g.CommentOption || []
          const option = pickRandom(options)
          return option ? resolveVars(option.text) : ''
        })
      } else if (groups.length > 0) {
        subjectParts = groups.map(g => {
          const options = g.CommentOption || []
          const option = pickRandom(options)
          return option ? resolveVars(option.text) : ''
        })
      } else {
        subjectParts = []
      }
      const subjectContent = subjectParts.filter(Boolean).join(' ')
      result = result.replaceAll('<SCG>', subjectContent || `[no subject groups for ${selectedSubject.code}]`)
    } else {
      result = result.replaceAll('<SCG>', '[no subject selected]')
    }

    // Resolve any remaining standard variables (e.g. in plain text portions of the template)
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

    return segments
  }, [template, commonGroups, selectedSubject, seed])

  return (
    <div>
      <div className="mb-6">
        <h2 className="text-lg font-bold text-gray-900 dark:text-white">Comment Format Template</h2>
        <p className="text-sm text-gray-500 mt-1">
          Define the global comment structure using group tags and {"<SCG>"} for subject-specific comment groups.
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Template Editor */}
        <div className="lg:col-span-2 space-y-4">
          <div className="bg-white dark:bg-gray-900 rounded-lg border border-gray-200 dark:border-gray-700 p-4">
            <label className="block text-sm font-medium text-gray-700 dark:text-gray-300 mb-2">
              Format Template
            </label>
            <textarea
              value={template}
              onChange={(e) => setTemplate(e.target.value)}
              rows={8}
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-800 dark:text-white shadow-sm border p-3 text-sm font-mono"
              placeholder={"<Academic>\n\n<Effort> <Behaviour>\n\n<SCG>\n\n<Overall>"}
            />
            <div className="mt-3 flex items-center gap-3">
              <button
                onClick={handleSave}
                disabled={saving}
                className="px-4 py-2 bg-indigo-600 text-white rounded-lg text-sm font-medium hover:bg-indigo-700 disabled:opacity-50 transition-colors"
              >
                {saving ? "Saving..." : "Save Template"}
              </button>
              {saved && <span className="text-sm text-green-600 font-medium">Saved!</span>}
              {error && <span className="text-sm text-red-600">{error}</span>}
            </div>
          </div>

          {/* Live Preview */}
          <div className="bg-white dark:bg-gray-900 rounded-lg border border-gray-200 dark:border-gray-700 p-4">
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-sm font-medium text-gray-700 dark:text-gray-300">Preview <span className="text-xs font-normal text-gray-400">(random options)</span></h3>
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setSeed(s => s + 1)}
                  className="px-2 py-1 text-xs font-medium text-indigo-600 hover:text-indigo-800 dark:text-indigo-400 dark:hover:text-indigo-300 border border-indigo-200 dark:border-indigo-800 rounded hover:bg-indigo-50 dark:hover:bg-indigo-900/30 transition-colors"
                >
                  Randomize
                </button>
                <select
                  value={previewSubjectId}
                  onChange={(e) => setPreviewSubjectId(e.target.value)}
                  className="rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-800 dark:text-white border px-2 py-1 text-sm"
                >
                  {subjects.map(s => (
                    <option key={s.id} value={s.id}>{s.code}</option>
                  ))}
                </select>
              </div>
            </div>
            <div className="bg-gray-50 dark:bg-gray-800 rounded-md p-4 text-sm whitespace-pre-wrap text-gray-700 dark:text-gray-300">
              {previewSegments.length > 0 ? (
                previewSegments.map((seg, i) =>
                  seg.unresolved ? (
                    <span key={i} className="text-red-500 font-medium">{seg.text}</span>
                  ) : (
                    <span key={i}>{seg.text}</span>
                  )
                )
              ) : (
                <span className="italic text-gray-400">Empty template</span>
              )}
            </div>
          </div>
        </div>

        {/* Sidebar */}
        <div className="space-y-4">
          {/* Available Tags */}
          <div className="bg-white dark:bg-gray-900 rounded-lg border border-gray-200 dark:border-gray-700 p-4">
            <h3 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">Available Tags</h3>
            <div className="space-y-2">
              {availableTags.map(tag => (
                <div key={tag} className="flex items-center gap-2">
                  <code className="text-xs bg-indigo-50 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-300 px-2 py-1 rounded font-mono">
                    {tag}
                  </code>
                </div>
              ))}
              <div className="flex items-center gap-2">
                <code className="text-xs bg-emerald-50 dark:bg-emerald-900/30 text-emerald-700 dark:text-emerald-300 px-2 py-1 rounded font-mono">
                  {"<SCG>"}
                </code>
                <span className="text-xs text-gray-500">Subject Comment Groups</span>
              </div>
            </div>
          </div>

          {/* Subject Formats */}
          <div className="bg-white dark:bg-gray-900 rounded-lg border border-gray-200 dark:border-gray-700 p-4">
            <h3 className="text-sm font-medium text-gray-700 dark:text-gray-300 mb-3">
              {"<SCG>"} Expansion
            </h3>
            <p className="text-xs text-gray-500 mb-3">
              Each subject defines how its comment groups combine when {"<SCG>"} is expanded.
            </p>
            <div className="space-y-2">
              {subjects.map((s: any) => (
                <div key={s.id} className="flex items-center justify-between">
                  <span className="text-xs font-medium text-gray-700 dark:text-gray-300">{s.code}</span>
                  <code className="text-xs bg-gray-100 dark:bg-gray-800 text-gray-600 dark:text-gray-400 px-2 py-0.5 rounded font-mono">
                    {s.commentFormat || "default order"}
                  </code>
                </div>
              ))}
              {subjects.length === 0 && (
                <p className="text-xs text-gray-400 italic">No subjects configured</p>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
