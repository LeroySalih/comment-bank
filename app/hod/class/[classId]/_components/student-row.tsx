"use client"

import { useState } from "react"
import { updateStudent, deleteStudent } from "@/lib/server-actions/hod"
import { useRouter } from "next/navigation"
import { Pencil, Trash2, X, Check } from "lucide-react"

interface StudentRowProps {
  student: {
    id: string
    Pupil: {
      admissionNumber: string
      firstName: string
      lastName: string
      gender: string
    }
    targetLevel: string | null
    actualLevel: string | null
  }
  classId: string
}

export function StudentRow({ student, classId }: StudentRowProps) {
  const [isEditing, setIsEditing] = useState(false)
  const [isDeleting, setIsDeleting] = useState(false)
  const [error, setError] = useState<string | null>(null)
  
  // Form state
  const [firstName, setFirstName] = useState(student.Pupil.firstName)
  const [lastName, setLastName] = useState(student.Pupil.lastName)
  const [gender, setGender] = useState(student.Pupil.gender)
  const [targetLevel, setTargetLevel] = useState(student.targetLevel || "")
  const [actualLevel, setActualLevel] = useState(student.actualLevel || "")

  const router = useRouter()

  async function handleUpdate(e: React.FormEvent) {
    e.preventDefault()
    setError(null)

    const formData = new FormData()
    formData.append("firstName", firstName)
    formData.append("lastName", lastName)
    formData.append("gender", gender)
    formData.append("targetLevel", targetLevel)
    formData.append("actualLevel", actualLevel)

    const result = await updateStudent(student.id, classId, formData)
    if (result.success) {
      setIsEditing(false)
      router.refresh()
    } else {
      setError('error' in result ? result.error : "Failed to update student")
    }
  }

  function handleCancel() {
    setIsEditing(false)
    setFirstName(student.Pupil.firstName)
    setLastName(student.Pupil.lastName)
    setGender(student.Pupil.gender)
    setTargetLevel(student.targetLevel || "")
    setActualLevel(student.actualLevel || "")
    setError(null)
  }

  async function handleDelete() {
    if (!confirm(`Are you sure you want to delete the assignment for ${student.Pupil.firstName} ${student.Pupil.lastName}? This will only remove them from this class.`)) return

    setIsDeleting(true)
    const result = await deleteStudent(student.id, classId)
    if (result.success) {
      router.refresh()
    } else {
      setError('error' in result ? result.error : "Failed to delete student")
      setIsDeleting(false)
    }
  }

  if (isEditing) {
    return (
      <tr className="bg-blue-50">
        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900" colSpan={4}>
          <form onSubmit={handleUpdate} className="flex flex-col gap-2">
            <div className="flex gap-2 items-center flex-wrap">
              <input 
                type="text" 
                value={lastName}
                onChange={(e) => setLastName(e.target.value)}
                className="border p-1 rounded text-sm w-24"
                placeholder="Last Name"
                required
              />
              <input 
                type="text" 
                value={firstName}
                onChange={(e) => setFirstName(e.target.value)}
                className="border p-1 rounded text-sm w-24"
                placeholder="First Name"
                required
              />
              <select 
                value={gender}
                onChange={(e) => setGender(e.target.value)}
                className="border p-1 rounded text-sm"
              >
                <option value="Male">Male</option>
                <option value="Female">Female</option>
              </select>
               <input 
                type="text" 
                value={targetLevel}
                onChange={(e) => setTargetLevel(e.target.value)}
                className="border p-1 rounded text-sm w-16"
                placeholder="Tgt Lvl"
              />
              <input 
                type="text" 
                value={actualLevel}
                onChange={(e) => setActualLevel(e.target.value)}
                className="border p-1 rounded text-sm w-16"
                placeholder="Act Lvl"
              />
            </div>
            
            <div className="flex gap-2">
              <button 
                type="submit" 
                className="flex items-center gap-1 bg-green-600 text-white px-2 py-1 rounded text-xs hover:bg-green-700"
              >
                <Check size={12} /> Save
              </button>
              <button 
                type="button" 
                onClick={handleCancel}
                className="flex items-center gap-1 bg-gray-300 text-gray-800 px-2 py-1 rounded text-xs hover:bg-gray-400"
              >
                <X size={12} /> Cancel
              </button>
            </div>
            {error && <p className="text-red-500 text-xs">{error}</p>}
          </form>
        </td>
      </tr>
    )
  }

  return (
    <tr className="hover:bg-gray-50 group">
      <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900 flex items-center gap-2">
        {student.Pupil.lastName}, {student.Pupil.firstName}
        <div className="opacity-0 group-hover:opacity-100 transition-opacity flex gap-1 ml-2">
          <button 
            onClick={() => setIsEditing(true)} 
            className="text-gray-400 hover:text-blue-600 transition-colors"
            title="Edit"
          >
            <Pencil size={14} />
          </button>
          <button 
            onClick={handleDelete} 
            disabled={isDeleting}
            className="text-gray-400 hover:text-red-600 transition-colors"
            title="Delete"
          >
            <Trash2 size={14} />
          </button>
        </div>
      </td>
      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
        {student.Pupil.gender}
      </td>
      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
        {student.targetLevel || "-"}
      </td>
      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
        {student.actualLevel || "-"}
      </td>
       {error && <td className="text-red-500 text-xs">{error}</td>}
    </tr>
  )
}
