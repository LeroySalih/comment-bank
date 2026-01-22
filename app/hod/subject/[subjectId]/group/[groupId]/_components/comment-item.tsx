"use client"

import { useState } from "react"
import { updateComment, deleteComment } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"
import { VariablePreview } from "@/components/VariablePreview"
import { countWords } from "@/lib/utils"

type CommentOption = {
  id: string
  code: string
  text: string
}

interface Props {
  comment: CommentOption
  subjectId: string
  groupId: string
}

export function CommentItem({ comment, subjectId, groupId }: Props) {
  const [isEditing, setIsEditing] = useState(false)
  const [loading, setLoading] = useState(false)
  const [text, setText] = useState(comment.text)
  const router = useRouter()

  async function handleUpdate(formData: FormData) {
    setLoading(true)
    const result = await updateComment(comment.id, subjectId, groupId, formData)
    setLoading(false)
    if (result.success) {
      setIsEditing(false)
      router.refresh()
    } else {
      alert("Failed to update comment")
    }
  }

  async function handleDelete() {
    if (!confirm("Are you sure you want to delete this comment?")) return
    setLoading(true)
    const result = await deleteComment(comment.id, subjectId, groupId)
    setLoading(false)
    if (result.success) {
      router.refresh()
    } else {
      alert("Failed to delete comment")
    }
  }

  if (isEditing) {
    return (
      <div className="p-4 border rounded-lg bg-indigo-50 border-indigo-200">
        <form action={handleUpdate} className="flex flex-col gap-3">
          <div className="flex gap-2">
            <input 
              type="text" 
              name="code" 
              defaultValue={comment.code}
              required 
              className="block w-20 rounded-md border-gray-300 shadow-sm border p-1 text-sm font-bold h-10"
            />
            <div className="flex-grow">
              <div className="flex justify-between items-end mb-1">
                <span className="text-[10px] text-gray-400 font-medium">{countWords(text)} words</span>
              </div>
              <textarea 
                name="text" 
                value={text}
                onChange={(e) => setText(e.target.value)}
                required 
                rows={2}
                className="block w-full rounded-md border-gray-300 shadow-sm border p-1 text-sm"
              />
            </div>
          </div>
          
          <VariablePreview text={text} />

          <div className="flex gap-2 justify-end">
            <button 
              type="submit" 
              disabled={loading}
              className="bg-indigo-600 text-white px-4 py-2 text-sm rounded-md hover:bg-indigo-700 disabled:opacity-50"
            >
              Save Changes
            </button>
            <button 
              type="button" 
              onClick={() => setIsEditing(false)}
              className="bg-gray-200 text-gray-800 px-4 py-2 text-sm rounded-md hover:bg-gray-300"
            >
              Cancel
            </button>
          </div>
        </form>
      </div>
    )
  }

  return (
    <div className="p-4 border rounded-lg hover:bg-gray-50 flex gap-4 group items-center">
      <div className="flex-shrink-0 w-16 font-bold text-gray-900 bg-gray-100 rounded flex items-center justify-center h-10 text-sm">
        {comment.code}
      </div>
      <div className="flex-grow">
        <p className="text-gray-800 text-sm">{comment.text}</p>
        <p className="text-[10px] text-gray-400 mt-1">{countWords(comment.text)} words</p>
      </div>
      <div className="opacity-0 group-hover:opacity-100 flex gap-2 transition-opacity">
        <button 
          onClick={() => {
            setText(comment.text)
            setIsEditing(true)
          }}
          className="text-indigo-600 hover:text-indigo-900 text-sm font-medium"
        >
          Edit
        </button>
        <button 
          onClick={handleDelete}
          disabled={loading}
          className="text-red-600 hover:text-red-900 text-sm font-medium disabled:opacity-50"
        >
          Delete
        </button>
      </div>
    </div>
  )
}
