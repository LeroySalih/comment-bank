"use client"

import { useState } from "react"
import { createStudent } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"

export function PupilForm({ classId }: { classId: string }) {
  const [error, setError] = useState<string | null>(null)
  const [open, setOpen] = useState(false)
  const router = useRouter()

  async function handleSubmit(formData: FormData) {
    setError(null)
    const result = await createStudent(classId, formData)
    if (!result.success) {
      setError(('error' in result ? result.error : "Failed to create student") || "Failed to create student")
    } else {
      setOpen(false)
      // reset form manually or just refresh
      router.refresh()
    }
  }

  if (!open) {
    return (
      <button 
        onClick={() => setOpen(true)}
        className="bg-indigo-600 text-white px-3 py-1 rounded-md text-sm hover:bg-indigo-700"
      >
        + Add Pupil
      </button>
    )
  }

  return (
    <div className="mt-4 border rounded p-4 bg-gray-50 mb-6">
      <h3 className="text-sm font-medium mb-3">Add Pupil</h3>
      <form action={handleSubmit} className="flex flex-col gap-3">
        <div className="grid grid-cols-3 gap-3">
          <div>
            <label className="block text-xs font-medium text-gray-500">Adm. Number</label>
            <input 
              type="text" 
              name="admissionNumber" 
              required 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm font-mono"
              placeholder="e.g. 123456"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500">First Name</label>
            <input 
              type="text" 
              name="firstName" 
              required 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500">Last Name</label>
            <input 
              type="text" 
              name="lastName" 
              required 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
            />
          </div>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <div>
            <label className="block text-xs font-medium text-gray-500">Gender</label>
            <select 
              name="gender" 
              required 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
            >
              <option value="Male">Male</option>
              <option value="Female">Female</option>
            </select>
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500">Target Level</label>
            <input 
              type="text" 
              name="targetLevel" 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
              placeholder="e.g. 5"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-500">Actual Level</label>
            <input 
              type="text" 
              name="actualLevel" 
              className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
              placeholder="e.g. 4"
            />
          </div>
        </div>
        <div className="flex gap-2 mt-2">
           <button 
            type="submit" 
            className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700"
          >
            Add Pupil
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
