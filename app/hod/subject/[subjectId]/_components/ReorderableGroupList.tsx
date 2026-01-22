"use client"

import { useState, useEffect } from "react"
import { DragDropContext, Droppable, Draggable, DropResult } from "@hello-pangea/dnd"
import { reorderCommentGroups } from "@/lib/server-actions/hod"
import { EditGroupForm } from "./edit-group-form"
import Link from "next/link"
import { GripVertical } from "lucide-react"

interface Group {
  id: string
  name: string
  _count: {
    options: number
  }
}

interface Props {
  subjectId: string
  initialGroups: Group[]
}

export function ReorderableGroupList({ subjectId, initialGroups }: Props) {
  const [groups, setGroups] = useState(initialGroups)
  
  useEffect(() => {
    setGroups(initialGroups)
  }, [initialGroups])

  const onDragEnd = async (result: DropResult) => {
    if (!result.destination) return

    const items = Array.from(groups)
    const [reorderedItem] = items.splice(result.source.index, 1)
    items.splice(result.destination.index, 0, reorderedItem)

    setGroups(items)

    // Update orders in DB
    const updates = items.map((group, index) => ({
      id: group.id,
      order: index, // The server action still takes 'order' as a parameter name for index, but updates 'displayOrder' field
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
              <Draggable key={group.id} draggableId={group.id} index={index}>
                {(provided) => (
                  <div
                    ref={provided.innerRef}
                    {...provided.draggableProps}
                    className="flex items-center gap-2"
                  >
                    <div 
                      {...provided.dragHandleProps}
                      className="text-gray-400 hover:text-gray-600 cursor-grab active:cursor-grabbing p-1"
                    >
                      <GripVertical size={20} />
                    </div>
                    <Link 
                      href={`/hod/subject/${subjectId}/group/${group.id}`}
                      className="flex-1 flex justify-between items-center p-3 bg-gray-50 rounded border hover:border-indigo-500 transition-colors group/item"
                    >
                      <div className="flex items-center gap-3">
                        <span className="font-medium">{group.name}</span>
                        <div onClick={(e) => e.preventDefault()}>
                            <EditGroupForm 
                                groupId={group.id} 
                                subjectId={subjectId} 
                                initialName={group.name} 
                            />
                        </div>
                      </div>
                      <span className="text-sm text-gray-500">{group._count.options} Comments →</span>
                    </Link>
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
