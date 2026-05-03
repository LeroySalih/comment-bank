'use client'

import { EditCommentWithSpag } from '@/components/EditCommentWithSpag'
import { updateComment } from '@/lib/server-actions/hod'

interface CommentOption {
  id: string
  code: string
  text: string
  displayOrder: number
}

interface EditCommentFormProps {
  comment: CommentOption
  subjectId: string
  groupId: string
  subjectTitle: string
  onClose: () => void
}

export function EditCommentForm({
  comment,
  subjectId,
  groupId,
  subjectTitle,
  onClose,
}: EditCommentFormProps) {
  async function handleSave(code: string, text: string) {
    const formData = new FormData()
    formData.append('code', code)
    formData.append('text', text)
    return updateComment(comment.id, subjectId, groupId, formData)
  }

  return (
    <EditCommentWithSpag
      code={comment.code}
      text={comment.text}
      subjectTitle={subjectTitle}
      onSave={handleSave}
      onClose={onClose}
    />
  )
}
