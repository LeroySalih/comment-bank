'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import { X } from 'lucide-react'
import { InlineSpagEditor } from '@/components/InlineSpagEditor'
import { VariablePreview } from '@/components/VariablePreview'
import { countWords } from '@/lib/utils'

interface EditCommentWithSpagProps {
  code: string
  text: string
  subjectTitle?: string
  onSave: (code: string, text: string) => Promise<{ success: boolean; error?: string }>
  onClose: () => void
}

export function EditCommentWithSpag({
  code: initialCode,
  text: initialText,
  subjectTitle = '',
  onSave,
  onClose,
}: EditCommentWithSpagProps) {
  const [code, setCode] = useState(initialCode)
  const [text, setText] = useState(initialText)
  const [saving, setSaving] = useState(false)
  const [saveError, setSaveError] = useState<string | null>(null)
  const router = useRouter()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setSaveError(null)
    setSaving(true)
    const result = await onSave(code, text)
    setSaving(false)
    if (!result.success) {
      setSaveError(result.error || 'Failed to save')
    } else {
      onClose()
      router.refresh()
    }
  }

  return (
    <div
      className="p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-200 dark:border-blue-800"
      onClick={(e) => e.stopPropagation()}
    >
      <div className="flex justify-between items-center mb-3">
        <h4 className="text-sm font-medium text-gray-900 dark:text-white">Edit Comment</h4>
        <button type="button" onClick={onClose} className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300">
          <X size={16} />
        </button>
      </div>

      <form onSubmit={handleSubmit} className="space-y-3">
        <div className="flex gap-3">
          {/* Code field */}
          <div className="w-24 shrink-0">
            <label className="block text-xs font-medium text-gray-500 dark:text-gray-400 mb-1">Code</label>
            <input
              type="text"
              value={code}
              onChange={(e) => setCode(e.target.value)}
              required
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>

          {/* Text field with inline SPAG */}
          <div className="flex-1 min-w-0">
            <div className="flex justify-between items-end mb-1">
              <label className="block text-xs font-medium text-gray-500 dark:text-gray-400">Comment Text</label>
              <span className="text-[10px] text-gray-400 font-medium">{countWords(text)} words</span>
            </div>
            <InlineSpagEditor
              value={text}
              onChange={setText}
              subjectTitle={subjectTitle}
              rows={3}
              placeholder="Write your comment…"
              spagOnly
            />
          </div>
        </div>

        <VariablePreview text={text} subjectName={subjectTitle || undefined} />

        {saveError && (
          <p className="text-xs text-red-600 dark:text-red-400">{saveError}</p>
        )}

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
            disabled={saving || !text.trim() || !code.trim()}
            className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {saving ? 'Saving…' : 'Save Changes'}
          </button>
        </div>
      </form>
    </div>
  )
}
