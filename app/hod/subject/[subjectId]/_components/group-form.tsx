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
      setError(result.error || "Failed to create group")
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
    <div className="mt-4 border rounded p-4 bg-gray-50">
      <h3 className="text-sm font-medium mb-2">Add Comment Group</h3>
      <form action={handleSubmit} className="flex gap-2 items-end">
        <div className="flex-1">
          <label className="block text-xs font-medium text-gray-500">Group Name (e.g. Work Ethic)</label>
          <input 
            type="text" 
            name="name" 
            required 
            className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
          />
        </div>
        <button 
          type="submit" 
          className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700"
        >
          Add
        </button>
        <button 
          type="button" 
          onClick={() => setOpen(false)}
          className="bg-gray-200 text-gray-800 px-3 py-1.5 rounded-md text-sm hover:bg-gray-300"
        >
          Cancel
        </button>
      </form>
      {error && <p className="text-red-600 mt-1 text-xs">{error}</p>}
    </div>
  )
}
