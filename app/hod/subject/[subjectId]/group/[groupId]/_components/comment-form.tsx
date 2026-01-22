"use client"

import { useState } from "react"
import { createComment } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"
import { VariablePreview } from "@/components/VariablePreview"
import { countWords } from "@/lib/utils"

export function CommentForm({ groupId, subjectId }: { groupId: string, subjectId: string }) {
  const [error, setError] = useState<string | null>(null)
  const [open, setOpen] = useState(false)
  const [text, setText] = useState("")
  const router = useRouter()

  async function handleSubmit(formData: FormData) {
    setError(null)
    const result = await createComment(groupId, subjectId, formData)
    if (!result.success) {
      setError(result.error || "Failed to create comment")
    } else {
      setOpen(false)
      setText("")
      router.refresh()
    }
  }

  if (!open) {
    return (
      <button 
        onClick={() => setOpen(true)}
        className="bg-indigo-600 text-white px-3 py-1 rounded-md text-sm hover:bg-indigo-700"
      >
        + Add Comment
      </button>
    )
  }

  return (
    <div className="mt-4 border rounded p-4 bg-gray-50 mb-6">
      <h3 className="text-sm font-medium mb-3">Add Comment</h3>
      <form action={handleSubmit} className="flex flex-col gap-3">
        <div className="flex gap-4">
          <div className="w-1/4">
            <label className="block text-xs font-medium text-gray-500">Code (e.g. H1, M2)</label>
            <input 
              type="text" 
              name="code" 
              required 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
            />
          </div>
          <div className="flex-grow">
            <div className="flex justify-between items-end">
              <label className="block text-xs font-medium text-gray-500">Comment Text</label>
              <span className="text-[10px] text-gray-400 font-medium">{countWords(text)} words</span>
            </div>
            <textarea 
              name="text" 
              required 
              rows={2}
              value={text}
              onChange={(e) => setText(e.target.value)}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
              placeholder="e.g. <Name> has made excellent progress this term. <He> is a pleasure to teach."
            />
          </div>
        </div>

        <div className="px-1">
          <VariablePreview text={text} />
        </div>

        <div className="flex gap-2">
           <button 
            type="submit" 
            className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700"
          >
            Add Comment
          </button>
          <button 
            type="button" 
            onClick={() => setOpen(false)}
            className="bg-gray-200 text-gray-800 px-3 py-1.5 rounded-md text-sm hover:bg-gray-300"
          >
            Cancel
          </button>
        </div>
      </form>
      {error && <p className="text-red-600 mt-1 text-xs">{error}</p>}
    </div>
  )
}
