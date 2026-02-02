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
      <div className="mt-4 border-t pt-4" onClick={(e) => e.stopPropagation()}>
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
              rows={2}
            />
          </div>

          <VariablePreview text={introduction} subjectName={title || code} />

          <div className="flex gap-2">
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
    )
  }

  return (
    <div className="absolute top-4 right-4 flex gap-1 transition-opacity opacity-0 group-hover:opacity-100" onClick={(e) => e.stopPropagation()}>
      <button 
        onClick={() => setIsEditing(true)}
        className="p-2 text-gray-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-full transition-all"
        title="Edit Subject"
      >
        <Pencil size={18} />
      </button>
      <button 
        onClick={handleDelete}
        disabled={isDeleting}
        className="p-2 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded-full transition-all"
        title="Delete Subject"
      >
        <Trash2 size={18} />
      </button>
      {error && <p className="absolute top-full right-0 text-red-600 text-[10px] mt-1 whitespace-nowrap">{error}</p>}
    </div>
  )
}
