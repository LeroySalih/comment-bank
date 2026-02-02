"use client"

import { useState } from "react"
import { createCommentGroup } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"

export function GroupForm({ subjectId }: { subjectId: string }) {
  const [error, setError] = useState<string | null>(null)
  const [open, setOpen] = useState(false)
  const router = useRouter()

  async function handleSubmit(formData: FormData) {
    setError(null)
    const result = await createCommentGroup(subjectId, formData)
    if (!result.success) {
      setError(('error' in result ? result.error : "Failed to create group") || "Failed to create group")
    } else {
      setOpen(false)
      router.refresh()
    }
  }

  if (!open) {
    return (
      <button 
        onClick={() => setOpen(true)}
        className="bg-indigo-600 text-white px-3 py-1 rounded-md text-sm hover:bg-indigo-700"
      >
        + Add Group
      </button>
    )
  }

  return (
    <div className="mt-4 border rounded p-4 bg-gray-50 dark:bg-gray-800 dark:border-gray-700">
      <h3 className="text-sm font-medium mb-3 text-gray-900 dark:text-white">Add Comment Group</h3>
      <form action={handleSubmit} className="space-y-3">
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block text-xs font-medium text-gray-500 dark:text-gray-400 mb-1">Code</label>
            <input
              type="text"
              name="name"
              required
              placeholder="e.g. WE"
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500 dark:text-gray-400 mb-1">Title</label>
            <input
              type="text"
              name="title"
              required
              placeholder="e.g. Work Ethic"
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
        </div>
        <div className="flex gap-2 justify-end">
          <button
            type="button"
            onClick={() => setOpen(false)}
            className="bg-gray-200 dark:bg-gray-600 text-gray-800 dark:text-gray-200 px-3 py-1.5 rounded-md text-sm hover:bg-gray-300 dark:hover:bg-gray-500"
          >
            Cancel
          </button>
          <button
            type="submit"
            className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700"
          >
            Add Group
          </button>
        </div>
      </form>
      {error && <p className="text-red-600 mt-2 text-xs">{error}</p>}
    </div>
  )
}
