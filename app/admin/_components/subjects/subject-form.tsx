"use client"

import { useState } from "react"
import { createSubject } from "@/lib/server-actions/admin"
import { useRouter } from "next/navigation"
import { VariablePreview } from "@/components/VariablePreview"

export function SubjectForm() {
  const [error, setError] = useState<string | null>(null)
  const [open, setOpen] = useState(false)
  const [code, setCode] = useState("")
  const [title, setTitle] = useState("")
  const [introduction, setIntroduction] = useState("")
  const router = useRouter()
  
  async function handleSubmit(formData: FormData) {
    setError(null)
    const result = await createSubject(formData)
    if (!result.success) {
      setError(result.error || "Failed to create subject")
    } else {
      setOpen(false)
      setCode("")
      setTitle("")
      setIntroduction("")
      router.refresh()
    }
  }

  if (!open) {
    return (
      <button 
        onClick={() => setOpen(true)}
        className="bg-indigo-600 text-white px-4 py-2 rounded-md hover:bg-indigo-700"
      >
        Add Subject
      </button>
    )
  }

  return (
    <div className="bg-white shadow rounded-lg p-6 mb-8">
      <h2 className="text-xl font-semibold mb-4">Add New Subject</h2>
      <form action={handleSubmit} className="flex flex-col gap-4">
        <div className="flex gap-4">
          <div className="w-1/3">
            <label className="block text-sm font-medium text-gray-700">Code (e.g. 7CS)</label>
            <input 
              type="text" 
              name="code" 
              value={code}
              onChange={(e) => setCode(e.target.value)}
              required 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2"
            />
          </div>
          <div className="w-2/3">
            <label className="block text-sm font-medium text-gray-700">Title (e.g. Year 7 Computer Science)</label>
            <input 
              type="text" 
              name="title" 
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2"
            />
          </div>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Introduction/Comment</label>
          <textarea 
            name="introduction" 
            value={introduction}
            onChange={(e) => setIntroduction(e.target.value)}
            className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2"
            rows={3}
          />
        </div>

        <VariablePreview text={introduction} subjectName={title || code} />

        <div className="flex gap-2">
          <button 
            type="submit" 
            className="bg-indigo-600 text-white px-4 py-2 rounded-md hover:bg-indigo-700"
          >
            Create
          </button>
          <button 
            type="button" 
            onClick={() => setOpen(false)}
            className="bg-gray-200 text-gray-800 px-4 py-2 rounded-md hover:bg-gray-300"
          >
            Cancel
          </button>
        </div>
      </form>
      {error && <p className="text-red-600 mt-2 text-sm">{error}</p>}
    </div>
  )
}
