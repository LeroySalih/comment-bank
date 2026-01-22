"use client"

import { useState, useEffect } from "react"
import { DragDropContext, Droppable, Draggable, DropResult } from "@hello-pangea/dnd"
import { CommentItem } from "./comment-item"
import { reorderComments } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"

type CommentOption = {
  id: string
  code: string
  text: string
  order: number | null
}

interface Props {
  initialComments: CommentOption[]
  subjectId: string
  groupId: string
}

export function CommentList({ initialComments, subjectId, groupId }: Props) {
  const [comments, setComments] = useState(initialComments)
  const router = useRouter()

  useEffect(() => {
    setComments(initialComments)
  }, [initialComments])

  async function onDragEnd(result: DropResult) {
    if (!result.destination) {
      return
    }

    const sourceIndex = result.source.index
    const destinationIndex = result.destination.index

    if (sourceIndex === destinationIndex) {
      return
    }

    const newComments = Array.from(comments)
    const [movedComment] = newComments.splice(sourceIndex, 1)
    newComments.splice(destinationIndex, 0, movedComment)

    setComments(newComments)

    const reorderedItems = newComments.map((comment, index) => ({
      id: comment.id,
      order: index
    }))

    await reorderComments(groupId, reorderedItems)
    router.refresh()
  }

  return (
    <DragDropContext onDragEnd={onDragEnd}>
      <Droppable droppableId="comments">
        {(provided) => (
          <div 
            {...provided.droppableProps} 
            ref={provided.innerRef}
            className="flex flex-col gap-4"
          >
            {comments.map((comment, index) => (
              <Draggable key={comment.id} draggableId={comment.id} index={index}>
                {(provided) => (
                   <div 
                     ref={provided.innerRef}
                     {...provided.draggableProps}
                     {...provided.dragHandleProps}
                   >
                     <CommentItem 
                        comment={comment} 
                        subjectId={subjectId} 
                        groupId={groupId} 
                      />
                   </div>
                )}
              </Draggable>
            ))}
            {provided.placeholder}
          </div>
        )}
      </Droppable>
    </DragDropContext>
  )
}
