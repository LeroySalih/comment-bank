"use client"

import { useState } from "react"
import { updateComment, deleteComment } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"
import { countWords } from "@/lib/utils"
import { EditCommentWithSpag } from "@/components/EditCommentWithSpag"

type CommentOption = {
  id: string
  code: string
  text: string
  subjectTitle?: string
}

interface Props {
  comment: CommentOption
  subjectId: string
  groupId: string
  subjectTitle?: string
}

export function CommentItem({ comment, subjectId, groupId, subjectTitle }: Props) {
  const [isEditing, setIsEditing] = useState(false)
  const [loading, setLoading] = useState(false)
  const router = useRouter()

  async function handleSave(code: string, text: string) {
    setLoading(true)
    const formData = new FormData()
    formData.append("code", code)
    formData.append("text", text)
    const result = await updateComment(comment.id, subjectId, groupId, formData)
    setLoading(false)
    return result
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
      <EditCommentWithSpag
        code={comment.code}
        text={comment.text}
        subjectTitle={subjectTitle}
        onSave={handleSave}
        onClose={() => setIsEditing(false)}
      />
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
          onClick={() => setIsEditing(true)}
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
