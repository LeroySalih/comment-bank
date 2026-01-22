"use client"

import { useState, useRef, useEffect } from "react"
import { assignTeachersToClass } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"
import { Users, Check, ChevronDown, X } from "lucide-react"

interface Teacher {
  id: string
  username: string
}

interface TeacherAssignmentProps {
  classId: string
  currentTeachers: Teacher[]
  availableTeachers: Teacher[]
}

export function TeacherAssignment({ classId, currentTeachers, availableTeachers }: TeacherAssignmentProps) {
  const [isOpen, setIsOpen] = useState(false)
  const [selectedIds, setSelectedIds] = useState<string[]>(currentTeachers.map(t => t.id))
  const [isSaving, setIsSaving] = useState(false)
  const dropdownRef = useRef<HTMLDivElement>(null)
  const router = useRouter()

  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false)
      }
    }
    document.addEventListener("mousedown", handleClickOutside)
    return () => document.removeEventListener("mousedown", handleClickOutside)
  }, [])

  async function handleToggle(teacherId: string) {
    const newSelected = selectedIds.includes(teacherId)
      ? selectedIds.filter(id => id !== teacherId)
      : [...selectedIds, teacherId]
    
    setSelectedIds(newSelected)
  }

  async function handleSave() {
    setIsSaving(true)
    const result = await assignTeachersToClass(classId, selectedIds)
    setIsSaving(false)
    if (result.success) {
      setIsOpen(false)
      router.refresh()
    } else {
      alert(result.error || "Failed to save teachers")
    }
  }

  const currentTeacherNames = availableTeachers
    .filter(t => selectedIds.includes(t.id))
    .map(t => t.username)
    .join(", ")

  return (
    <div className="relative" ref={dropdownRef} onClick={(e) => e.stopPropagation()}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        className="flex items-center gap-2 px-3 py-1.5 bg-white border border-gray-300 rounded-md text-sm hover:bg-gray-50 transition-colors"
        title="Assign Teachers"
      >
        <Users size={16} className="text-gray-500" />
        <span className="max-w-[150px] truncate">
          {selectedIds.length === 0 ? "Assign Teachers" : currentTeacherNames}
        </span>
        <ChevronDown size={14} className="text-gray-400" />
      </button>

      {isOpen && (
        <div className="absolute z-10 mt-1 w-64 bg-white border border-gray-200 rounded-md shadow-lg overflow-hidden">
          <div className="p-2 border-b bg-gray-50 flex justify-between items-center">
            <span className="text-xs font-semibold uppercase text-gray-500">Edit Teachers</span>
            <button onClick={() => setIsOpen(false)} className="text-gray-400 hover:text-gray-600">
              <X size={14} />
            </button>
          </div>
          <div className="max-h-60 overflow-y-auto p-1">
            {availableTeachers.map((teacher) => (
              <label
                key={teacher.id}
                className="flex items-center gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer transition-colors"
                onClick={(e) => e.stopPropagation()}
              >
                <input
                  type="checkbox"
                  checked={selectedIds.includes(teacher.id)}
                  onChange={() => handleToggle(teacher.id)}
                  className="rounded border-gray-300 text-indigo-600 focus:ring-indigo-500"
                />
                <span className="text-sm text-gray-700">{teacher.username}</span>
                {selectedIds.includes(teacher.id) && <Check size={14} className="ml-auto text-indigo-600" />}
              </label>
            ))}
            {availableTeachers.length === 0 && (
              <p className="text-xs text-gray-500 p-2 italic text-center">No teachers available.</p>
            )}
          </div>
          <div className="p-2 border-t bg-gray-50 flex justify-end gap-2">
             <button
              onClick={() => {
                setSelectedIds(currentTeachers.map(t => t.id))
                setIsOpen(false)
              }}
              className="px-2 py-1 text-xs text-gray-600 hover:text-gray-800"
            >
              Cancel
            </button>
            <button
              onClick={handleSave}
              disabled={isSaving}
              className="px-3 py-1 bg-indigo-600 text-white rounded text-xs hover:bg-indigo-700 disabled:opacity-50 flex items-center gap-1"
            >
              {isSaving ? "Saving..." : "Save Changes"}
            </button>
          </div>
        </div>
      )}
    </div>
  )
}
