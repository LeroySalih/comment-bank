"use client"

import { useState } from "react"
import { updateCommentGroup, deleteCommentGroup } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"
import { Pencil, Trash2, X } from "lucide-react"

interface EditGroupFormProps {
  groupId: string
  subjectId: string
  initialName: string
}

export function EditGroupForm({ groupId, subjectId, initialName }: EditGroupFormProps) {
  const [isEditing, setIsEditing] = useState(false)
  const [name, setName] = useState(initialName)
  const [error, setError] = useState<string | null>(null)
  const [isDeleting, setIsDeleting] = useState(false)
  
  const router = useRouter()

  async function handleUpdate(e: React.FormEvent) {
    e.preventDefault()
    setError(null)
    
    const formData = new FormData()
    formData.append("name", name)

    const result = await updateCommentGroup(groupId, subjectId, formData)
    if (result.success) {
      setIsEditing(false)
      router.refresh()
    } else {
      setError(result.error || "Failed to update group")
    }
  }

  async function handleDelete() {
    if (!confirm("Are you sure you want to delete this group?")) return

    setIsDeleting(true)
    const result = await deleteCommentGroup(groupId, subjectId)
    if (result.success) {
      router.refresh()
    } else {
      setError(result.error || "Failed to delete group")
      setIsDeleting(false)
    }
  }

  if (isEditing) {
    return (
      <div className="flex items-center gap-2" onClick={(e) => e.stopPropagation()}>
        <form onSubmit={handleUpdate} className="flex items-center gap-2">
          <input
            type="text"
            value={name}
            onChange={(e) => setName(e.target.value)}
            className="border rounded px-2 py-1 text-sm w-32"
            required
            autoFocus
          />
          <button 
            type="submit"
            className="text-green-600 hover:text-green-800 p-1"
          >
            Save
          </button>
          <button 
            type="button" 
            onClick={() => {
              setIsEditing(false)
              setName(initialName)
              setError(null)
            }}
            className="text-gray-500 hover:text-gray-700 p-1"
          >
            <X size={16} />
          </button>
        </form>
        {error && <span className="text-red-500 text-xs">{error}</span>}
      </div>
    )
  }

  // Use a div to stop propagation when clicking actions
  return (
    <div className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
      <button 
        onClick={(e) => {
          e.preventDefault();
          setIsEditing(true);
        }}
        className="p-1 text-gray-500 hover:text-blue-600 transition-colors"
        title="Edit Group Name"
      >
        <Pencil size={14} />
      </button>
      <button 
        onClick={(e) => {
            e.preventDefault();
            handleDelete();
        }}
        disabled={isDeleting}
        className="p-1 text-gray-500 hover:text-red-600 transition-colors"
        title="Delete Group"
      >
        <Trash2 size={14} />
      </button>
       {error && <span className="text-red-500 text-xs">{error}</span>}
    </div>
  )
}
