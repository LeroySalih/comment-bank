"use client"

import { useState } from "react"
import { useRouter } from "next/navigation"
import { createUser } from "@/lib/server-actions/create-user"

export function CreateUserForm({ roles }: { roles: { id: string, name: string }[] }) {
  const [error, setError] = useState<string | null>(null)
  const router = useRouter() // Added for router.refresh()
  
  async function handleSubmit(formData: FormData) {
    setError(null)
    const result = await createUser(formData)
    if (!result.success) {
      setError(('error' in result ? result.error : "Failed to create user") || "Failed to create user")
    } else {
      // Optional: Reset form or show success toast
      alert("User created!")
      router.refresh()
    }
  }

  return (
    <div className="bg-white shadow rounded-lg p-6 mb-8">
      <h2 className="text-xl font-semibold mb-4">Add New User</h2>
      <form action={handleSubmit} className="flex flex-col gap-4 md:flex-row md:items-end">
        <div className="flex-1">
          <label className="block text-sm font-medium text-gray-700">Username</label>
          <input 
            type="text" 
            name="username" 
            required 
            className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2"
          />
        </div>
        <div className="flex-1">
          <label className="block text-sm font-medium text-gray-700">Password</label>
          <input 
            type="password" 
            name="password" 
            required 
            className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2"
          />
        </div>
        <div className="flex-1">
          <label className="block text-sm font-medium text-gray-700">Initial Role</label>
          <select 
            name="role" 
            className="mt-1 block w-full rounded-md border-gray-300 shadow-sm border p-2"
          >
            <option value="">None</option>
            {roles.map(r => (
              <option key={r.id} value={r.name}>{r.name}</option>
            ))}
          </select>
        </div>
        <div>
          <button 
            type="submit" 
            className="bg-indigo-600 text-white px-4 py-2 rounded-md hover:bg-indigo-700"
          >
            Create User
          </button>
        </div>
      </form>
      {error && <p className="text-red-600 mt-2 text-sm">{error}</p>}
    </div>
  )
}
