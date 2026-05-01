"use client"

import { useState } from "react"
import { updateSubject, deleteSubject } from "@/lib/server-actions/admin"
import { useRouter } from "next/navigation"
import { Pencil, Trash2, X, Check } from "lucide-react"
import { VariablePreview } from "@/components/VariablePreview"

interface EditSubjectFormProps {
  subject: {
    id: string
    code: string
    title: string | null
    studiedComment: string | null
  }
}

export function EditSubjectForm({ subject }: EditSubjectFormProps) {
  const [isEditing, setIsEditing] = useState(false)
  const [code, setCode] = useState(subject.code)
  const [title, setTitle] = useState(subject.title || "")
  const [introduction, setIntroduction] = useState(subject.studiedComment || "")
  const [error, setError] = useState<string | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)
  
  const router = useRouter()

  async function handleUpdate(e: React.FormEvent) {
    e.preventDefault()
    e.stopPropagation()
    setError(null)
    
    const formData = new FormData()
    formData.append("code", code)
    formData.append("title", title)
    formData.append("introduction", introduction)

    const result = await updateSubject(subject.id, formData)
    if (result.success) {
      setIsEditing(false)
      router.refresh()
    } else {
      setError('error' in result ? result.error : "Failed to update subject")
    }
  }


  async function handleDelete(e: React.MouseEvent) {
    e.preventDefault()
    e.stopPropagation()
    
    if (!confirm(`Are you sure you want to delete "${subject.code}"? This will delete all classes, pupils, groups and comments associated with it.`)) return

    setIsDeleting(true)
    const result = await deleteSubject(subject.id)
    if (result.success) {
      router.refresh()
    } else {
      setError(('error' in result ? result.error : "Failed to delete subject") || "Failed to delete subject")
      setIsDeleting(false)
    }
  }


  if (isEditing) {
    return (
      <div
        className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4"
        onClick={() => setIsEditing(false)}
      >
        <div
          className="bg-white dark:bg-gray-900 rounded-lg shadow-xl w-full max-w-xl p-6"
          onClick={(e) => e.stopPropagation()}
        >
          <h3 className="text-lg font-semibold mb-4 text-gray-900 dark:text-gray-100">Edit Subject</h3>
          <form onSubmit={handleUpdate} className="space-y-3">
            <div className="flex gap-2">
              <div className="w-1/3">
                <label className="block text-xs font-medium text-gray-500 uppercase">Code</label>
                <input
                  type="text"
                  value={code}
                  onChange={(e) => setCode(e.target.value)}
                  className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2 text-sm"
                  required
                  autoFocus
                />
              </div>
              <div className="w-2/3">
                <label className="block text-xs font-medium text-gray-500 uppercase">Title</label>
                <input
                  type="text"
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2 text-sm"
                />
              </div>
            </div>
            <div>
              <label className="block text-xs font-medium text-gray-500 uppercase">Introduction (Optional)</label>
              <textarea
                value={introduction}
                onChange={(e) => setIntroduction(e.target.value)}
                className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2 text-sm"
                rows={3}
              />
            </div>

            <VariablePreview text={introduction} subjectName={title || code} />

            <div className="flex gap-2 pt-2">
              <button
                type="submit"
                className="flex items-center gap-1 bg-green-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-green-700 transition-colors"
              >
                <Check size={16} /> Save Changes
              </button>
              <button
                type="button"
                onClick={() => setIsEditing(false)}
                className="flex items-center gap-1 bg-gray-100 text-gray-700 px-3 py-1.5 rounded-md text-sm hover:bg-gray-200 transition-colors"
              >
                <X size={16} /> Cancel
              </button>
            </div>
            {error && <p className="text-red-600 text-xs mt-1">{error}</p>}
          </form>
        </div>
      </div>
    )
  }

  return (
    <div className="inline-flex gap-1" onClick={(e) => e.stopPropagation()}>
      <button
        onClick={() => setIsEditing(true)}
        className="p-2 text-gray-500 dark:text-gray-400 hover:text-indigo-600 hover:bg-indigo-50 dark:hover:bg-indigo-900/30 rounded transition-colors"
        title="Edit subject"
        aria-label="Edit subject"
      >
        <Pencil size={16} />
      </button>
      <button
        onClick={handleDelete}
        disabled={isDeleting}
        className="p-2 text-gray-500 dark:text-gray-400 hover:text-red-600 hover:bg-red-50 dark:hover:bg-red-900/30 rounded transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
        title="Delete subject"
        aria-label="Delete subject"
      >
        <Trash2 size={16} />
      </button>
      {error && <span className="text-red-600 text-xs ml-2 self-center">{error}</span>}
    </div>
  )
}
