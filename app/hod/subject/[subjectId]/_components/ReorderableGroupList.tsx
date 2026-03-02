"use client"

import { useState, useEffect } from "react"
import { DragDropContext, Droppable, Draggable, DropResult } from "@hello-pangea/dnd"
import { reorderCommentGroups, createComment, updateComment, deleteComment, reorderComments } from "@/lib/server-actions/hod"
import { EditGroupForm } from "./edit-group-form"
import { GripVertical, Plus, X, ChevronDown, ChevronRight, Pencil, Trash2 } from "lucide-react"
import { useRouter } from "next/navigation"
import { VariablePreview } from "@/components/VariablePreview"
import { countWords } from "@/lib/utils"

interface CommentOption {
  id: string
  code: string
  text: string
  displayOrder: number
}

interface Group {
  id: string
  name: string
  title: string
  CommentOption: CommentOption[]
  _count: {
    CommentOption: number
  }
}

type CcgGroup = { id: string; name: string; title: string }

interface Props {
  subjectId: string
  initialGroups: Group[]
  ccgGroups?: CcgGroup[]
}

function InlineCommentForm({ groupId, subjectId, onClose }: { groupId: string; subjectId: string; onClose: () => void }) {
  const [code, setCode] = useState("")
  const [text, setText] = useState("")
  const [error, setError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)
  const router = useRouter()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError(null)
    setSaving(true)

    const formData = new FormData()
    formData.append("code", code)
    formData.append("text", text)

    const result = await createComment(groupId, subjectId, formData)
    setSaving(false)

    if (!result.success) {
      setError(('error' in result ? result.error : "Failed to create comment") || "Failed to create comment")
    } else {
      setCode("")
      setText("")
      onClose()
      router.refresh()
    }
  }

  return (
    <div className="mt-3 p-4 bg-gray-50 dark:bg-gray-800 rounded-lg border border-gray-200 dark:border-gray-700" onClick={(e) => e.stopPropagation()}>
      <div className="flex justify-between items-center mb-3">
        <h4 className="text-sm font-medium text-gray-900 dark:text-white">Add Comment</h4>
        <button onClick={onClose} className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300">
          <X size={16} />
        </button>
      </div>
      <form onSubmit={handleSubmit} className="space-y-3">
        <div className="flex gap-3">
          <div className="w-24">
            <label className="block text-xs font-medium text-gray-500 dark:text-gray-400 mb-1">Code</label>
            <input
              type="text"
              value={code}
              onChange={(e) => setCode(e.target.value)}
              required
              placeholder="e.g. H1"
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
          <div className="flex-1">
            <div className="flex justify-between items-end mb-1">
              <label className="block text-xs font-medium text-gray-500 dark:text-gray-400">Comment Text</label>
              <span className="text-[10px] text-gray-400 font-medium">{countWords(text)} words</span>
            </div>
            <textarea
              value={text}
              onChange={(e) => setText(e.target.value)}
              required
              rows={2}
              placeholder="e.g. <Name> has made excellent progress. <He> is a pleasure to teach."
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
        </div>

        <VariablePreview text={text} />

        <div className="flex gap-2 justify-end">
          <button
            type="button"
            onClick={onClose}
            className="bg-gray-200 dark:bg-gray-600 text-gray-800 dark:text-gray-200 px-3 py-1.5 rounded-md text-sm hover:bg-gray-300 dark:hover:bg-gray-500"
          >
            Cancel
          </button>
          <button
            type="submit"
            disabled={saving}
            className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700 disabled:opacity-50"
          >
            {saving ? "Adding..." : "Add Comment"}
          </button>
        </div>
        {error && <p className="text-red-600 text-xs">{error}</p>}
      </form>
    </div>
  )
}

function EditCommentForm({ comment, subjectId, groupId, onClose }: { comment: CommentOption; subjectId: string; groupId: string; onClose: () => void }) {
  const [code, setCode] = useState(comment.code)
  const [text, setText] = useState(comment.text)
  const [error, setError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)
  const router = useRouter()

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault()
    setError(null)
    setSaving(true)

    const formData = new FormData()
    formData.append("code", code)
    formData.append("text", text)

    const result = await updateComment(comment.id, subjectId, groupId, formData)
    setSaving(false)

    if (!result.success) {
      setError(('error' in result ? result.error : "Failed to update comment") || "Failed to update comment")
    } else {
      onClose()
      router.refresh()
    }
  }

  return (
    <div className="p-4 bg-blue-50 dark:bg-blue-900/20 rounded-lg border border-blue-200 dark:border-blue-800" onClick={(e) => e.stopPropagation()}>
      <div className="flex justify-between items-center mb-3">
        <h4 className="text-sm font-medium text-gray-900 dark:text-white">Edit Comment</h4>
        <button onClick={onClose} className="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300">
          <X size={16} />
        </button>
      </div>
      <form onSubmit={handleSubmit} className="space-y-3">
        <div className="flex gap-3">
          <div className="w-24">
            <label className="block text-xs font-medium text-gray-500 dark:text-gray-400 mb-1">Code</label>
            <input
              type="text"
              value={code}
              onChange={(e) => setCode(e.target.value)}
              required
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
          <div className="flex-1">
            <div className="flex justify-between items-end mb-1">
              <label className="block text-xs font-medium text-gray-500 dark:text-gray-400">Comment Text</label>
              <span className="text-[10px] text-gray-400 font-medium">{countWords(text)} words</span>
            </div>
            <textarea
              value={text}
              onChange={(e) => setText(e.target.value)}
              required
              rows={2}
              className="block w-full rounded-md border-gray-300 dark:border-gray-600 dark:bg-gray-700 dark:text-white shadow-sm border p-2 text-sm"
            />
          </div>
        </div>

        <VariablePreview text={text} />

        <div className="flex gap-2 justify-end">
          <button
            type="button"
            onClick={onClose}
            className="bg-gray-200 dark:bg-gray-600 text-gray-800 dark:text-gray-200 px-3 py-1.5 rounded-md text-sm hover:bg-gray-300 dark:hover:bg-gray-500"
          >
            Cancel
          </button>
          <button
            type="submit"
            disabled={saving}
            className="bg-indigo-600 text-white px-3 py-1.5 rounded-md text-sm hover:bg-indigo-700 disabled:opacity-50"
          >
            {saving ? "Saving..." : "Save Changes"}
          </button>
        </div>
        {error && <p className="text-red-600 text-xs">{error}</p>}
      </form>
    </div>
  )
}

function CommentItem({ comment, subjectId, groupId, index, onReorder }: {
  comment: CommentOption
  subjectId: string
  groupId: string
  index: number
  onReorder: (comments: CommentOption[]) => void
}) {
  const [isEditing, setIsEditing] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)
  const router = useRouter()

  async function handleDelete() {
    if (!confirm("Are you sure you want to delete this comment?")) return
    setIsDeleting(true)
    const result = await deleteComment(comment.id, subjectId, groupId)
    if (result.success) {
      router.refresh()
    } else {
      alert(('error' in result ? result.error : "Failed to delete comment") || "Failed to delete")
      setIsDeleting(false)
    }
  }

  if (isEditing) {
    return (
      <Draggable draggableId={comment.id} index={index}>
        {(provided) => (
          <div
            ref={provided.innerRef}
            {...provided.draggableProps}
            className="mb-2"
          >
            <EditCommentForm comment={comment} subjectId={subjectId} groupId={groupId} onClose={() => setIsEditing(false)} />
          </div>
        )}
      </Draggable>
    )
  }

  return (
    <Draggable draggableId={comment.id} index={index}>
      {(provided) => (
        <div
          ref={provided.innerRef}
          {...provided.draggableProps}
          className="flex items-start gap-2 p-2 bg-white dark:bg-gray-900 rounded border border-gray-100 dark:border-gray-700 hover:border-gray-300 dark:hover:border-gray-600 mb-2 group"
        >
          <div
            {...provided.dragHandleProps}
            className="text-gray-400 hover:text-gray-600 cursor-grab active:cursor-grabbing mt-1"
          >
            <GripVertical size={14} />
          </div>
          <div className="flex-1 min-w-0">
            <div className="flex items-center gap-2 mb-1">
              <span className="bg-indigo-100 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-300 text-xs font-bold px-1.5 py-0.5 rounded">
                {comment.code}
              </span>
              <span className="text-[10px] text-gray-400">{countWords(comment.text)} words</span>
            </div>
            <p className="text-sm text-gray-700 dark:text-gray-300 line-clamp-2">{comment.text}</p>
          </div>
          <div className="flex items-center gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
            <button
              onClick={() => setIsEditing(true)}
              className="p-1 text-gray-400 hover:text-blue-600 transition-colors"
              title="Edit"
            >
              <Pencil size={14} />
            </button>
            <button
              onClick={handleDelete}
              disabled={isDeleting}
              className="p-1 text-gray-400 hover:text-red-600 transition-colors disabled:opacity-50"
              title="Delete"
            >
              <Trash2 size={14} />
            </button>
          </div>
        </div>
      )}
    </Draggable>
  )
}

function CommentsList({ group, subjectId }: { group: Group; subjectId: string }) {
  const [comments, setComments] = useState(group.CommentOption)
  const router = useRouter()

  useEffect(() => {
    setComments(group.CommentOption)
  }, [group.CommentOption])

  const handleCommentDragEnd = async (result: DropResult) => {
    if (!result.destination) return

    const items = Array.from(comments)
    const [reorderedItem] = items.splice(result.source.index, 1)
    items.splice(result.destination.index, 0, reorderedItem)

    setComments(items)

    const updates = items.map((comment, index) => ({
      id: comment.id,
      order: index
    }))

    await reorderComments(group.id, updates)
    router.refresh()
  }

  return (
    <DragDropContext onDragEnd={handleCommentDragEnd}>
      <Droppable droppableId={`comments-${group.id}`}>
        {(provided) => (
          <div
            {...provided.droppableProps}
            ref={provided.innerRef}
          >
            {comments.map((comment, index) => (
              <CommentItem
                key={comment.id}
                comment={comment}
                subjectId={subjectId}
                groupId={group.id}
                index={index}
                onReorder={setComments}
              />
            ))}
            {provided.placeholder}
            {comments.length === 0 && (
              <p className="text-gray-400 text-sm italic py-2">No comments in this group yet.</p>
            )}
          </div>
        )}
      </Droppable>
    </DragDropContext>
  )
}

export function ReorderableGroupList({ subjectId, initialGroups, ccgGroups }: Props) {
  const [groups, setGroups] = useState(initialGroups)
  const [addingToGroup, setAddingToGroup] = useState<string | null>(null)
  const [expandedGroups, setExpandedGroups] = useState<Set<string>>(new Set())

  useEffect(() => {
    setGroups(initialGroups)
  }, [initialGroups])

  const toggleGroup = (groupId: string) => {
    setExpandedGroups(prev => {
      const next = new Set(prev)
      if (next.has(groupId)) {
        next.delete(groupId)
      } else {
        next.add(groupId)
      }
      return next
    })
  }

  const onDragEnd = async (result: DropResult) => {
    if (!result.destination) return
    if (!result.draggableId.startsWith('group-')) return

    const items = Array.from(groups)
    const [reorderedItem] = items.splice(result.source.index, 1)
    items.splice(result.destination.index, 0, reorderedItem)

    setGroups(items)

    const updates = items.map((group, index) => ({
      id: group.id,
      order: index
    }))

    await reorderCommentGroups(subjectId, updates)
  }

  return (
    <DragDropContext onDragEnd={onDragEnd}>
      <Droppable droppableId="groups">
        {(provided) => (
          <div
            {...provided.droppableProps}
            ref={provided.innerRef}
            className="space-y-4"
          >
            {groups.map((group, index) => (
              <Draggable key={group.id} draggableId={`group-${group.id}`} index={index}>
                {(provided) => (
                  <div
                    ref={provided.innerRef}
                    {...provided.draggableProps}
                    className="bg-white dark:bg-gray-900 rounded-lg border border-[#f0f2f4] dark:border-gray-700 overflow-hidden"
                  >
                    {/* Group Header */}
                    <div className="flex items-center gap-2 p-3 bg-gray-50 dark:bg-gray-800/50">
                      <div
                        {...provided.dragHandleProps}
                        className="text-[#617289] dark:text-gray-500 hover:text-primary cursor-grab active:cursor-grabbing"
                      >
                        <GripVertical size={20} />
                      </div>
                      <button
                        onClick={() => toggleGroup(group.id)}
                        className="text-gray-500 hover:text-gray-700 dark:hover:text-gray-300"
                      >
                        {expandedGroups.has(group.id) ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
                      </button>
                      <span className="bg-gray-200 dark:bg-gray-700 text-gray-600 dark:text-gray-300 text-xs font-bold px-2 py-1 rounded">
                        {group.name}
                      </span>
                      <span className="font-bold text-[#111418] dark:text-white flex-1">
                        {group.title}
                        {ccgGroups?.some(cg => cg.name === group.name) && (
                          <span className="text-[10px] font-medium text-indigo-600 bg-indigo-100 dark:bg-indigo-900/30 dark:text-indigo-300 px-1.5 py-0.5 rounded ml-1">
                            CCG Override
                          </span>
                        )}
                      </span>
                      <EditGroupForm
                        groupId={group.id}
                        subjectId={subjectId}
                        initialName={group.name}
                        initialTitle={group.title}
                      />
                      <button
                        onClick={() => {
                          if (!expandedGroups.has(group.id)) {
                            toggleGroup(group.id)
                          }
                          setAddingToGroup(addingToGroup === group.id ? null : group.id)
                        }}
                        className="text-xs font-medium text-indigo-600 hover:text-indigo-800 dark:text-indigo-400 dark:hover:text-indigo-300 flex items-center gap-1 ml-2"
                      >
                        <Plus size={14} />
                        Add
                      </button>
                      <span className="text-sm font-medium text-[#617289] dark:text-gray-400 ml-2">
                        {group._count.CommentOption} comments
                      </span>
                    </div>

                    {/* Expanded Content */}
                    {expandedGroups.has(group.id) && (
                      <div className="p-4 border-t border-gray-100 dark:border-gray-700">
                        {addingToGroup === group.id && (
                          <InlineCommentForm
                            groupId={group.id}
                            subjectId={subjectId}
                            onClose={() => setAddingToGroup(null)}
                          />
                        )}
                        <div className={addingToGroup === group.id ? "mt-4" : ""}>
                          <CommentsList group={group} subjectId={subjectId} />
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </Draggable>
            ))}
            {provided.placeholder}
            {groups.length === 0 && (
              <p className="text-gray-500 text-sm italic">No groups added yet.</p>
            )}
          </div>
        )}
      </Droppable>
    </DragDropContext>
  )
}
