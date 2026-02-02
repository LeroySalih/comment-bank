"use client"

import { useState } from "react"
import { updateDeadline, deleteDeadline } from "@/lib/server-actions/admin"
import { useRouter } from "next/navigation"

interface EditDeadlineFormProps {
  deadline: {
    id: string
    title: string
    date: string | Date
    description: string | null
    isActive: boolean
  }
}

export function EditDeadlineForm({ deadline }: EditDeadlineFormProps) {
  const [isEditing, setIsEditing] = useState(false)
  const [title, setTitle] = useState(deadline.title)
  const [date, setDate] = useState(
    new Date(deadline.date).toISOString().split('T')[0]
  )
  const [description, setDescription] = useState(deadline.description || "")
  const [isActive, setIsActive] = useState(deadline.isActive)
  const [error, setError] = useState<string | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)
  const router = useRouter()

  async function handleUpdate(e: React.FormEvent) {
    e.preventDefault()
    setError(null)

    const formData = new FormData()
    formData.append("title", title)
    formData.append("date", date)
    formData.append("description", description)
    formData.append("isActive", String(isActive))

    const result = await updateDeadline(deadline.id, formData)
    if (result.success) {
      setIsEditing(false)
      router.refresh()
    } else {
      setError('error' in result ? result.error : "Failed to update deadline")
    }
  }

  async function handleDelete() {
    if (!confirm(`Are you sure you want to delete "${deadline.title}"?`)) return

    setIsDeleting(true)
    const result = await deleteDeadline(deadline.id)
    if (result.success) {
      router.refresh()
    } else {
      setError('error' in result ? result.error : "Failed to delete deadline")
      setIsDeleting(false)
    }
  }

  if (isEditing) {
    return (
      <div className="mt-4 border-t pt-4 w-full">
        <form onSubmit={handleUpdate} className="space-y-3">
          <div className="flex gap-2">
            <div className="w-2/3">
              <label className="block text-xs font-medium text-gray-500 uppercase">Title</label>
              <input
                type="text"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2 text-sm"
                required
              />
            </div>
            <div className="w-1/3">
              <label className="block text-xs font-medium text-gray-500 uppercase">Date</label>
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2 text-sm"
                required
              />
            </div>
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 uppercase">Description</label>
            <textarea
              value={description}
              onChange={(e) => setDescription(e.target.value)}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2 text-sm"
              rows={2}
            />
          </div>
          <div className="flex items-center gap-2">
            <input
              type="checkbox"
              id={`isActive-${deadline.id}`}
              checked={isActive}
              onChange={(e) => setIsActive(e.target.checked)}
              className="rounded"
            />
            <label htmlFor={`isActive-${deadline.id}`} className="text-sm text-gray-700">Active</label>
          </div>
          <div className="flex gap-2">
            <button
              type="submit"
              className="flex items-center gap-1 bg-green-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-green-700"
            >
              Save
            </button>
            <button
              type="button"
              onClick={() => setIsEditing(false)}
              className="flex items-center gap-1 bg-gray-100 text-gray-700 px-3 py-1.5 rounded-md text-sm hover:bg-gray-200"
            >
              Cancel
            </button>
          </div>
          {error && <p className="text-red-600 text-xs mt-1">{error}</p>}
        </form>
      </div>
    )
  }

  return (
    <div className="flex gap-1 transition-opacity opacity-0 group-hover:opacity-100">
      <button
        onClick={() => setIsEditing(true)}
        className="p-2 text-gray-400 hover:text-indigo-600 hover:bg-indigo-50 rounded-full"
        title="Edit Deadline"
      >
        <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
        </svg>
      </button>
      <button
        onClick={handleDelete}
        disabled={isDeleting}
        className="p-2 text-gray-400 hover:text-red-600 hover:bg-red-50 rounded-full"
        title="Delete Deadline"
      >
        <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
        </svg>
      </button>
    </div>
  )
}
