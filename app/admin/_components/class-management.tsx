"use client"

import { useState, useEffect, useRef } from "react"
import { getClasses, createClassFromForm, deleteClass, getAllForms, getAllUsers, updateClass } from "@/lib/server-actions/admin"
import { EditClassSidebar } from "./edit-class-sidebar"

interface Class {
  id: string
  name: string
  year: string | null
  subjectId: string
  Subject: {
    code: string
    title: string | null
  }
  _count: {
    Assignment: number
  }
  User: Array<{
    id: string
    username: string
  }>
}

interface Subject {
  id: string
  code: string
  title: string | null
}

interface User {
  id: string
  username: string
}

interface ClassManagementProps {
  subjects: Subject[]
}

function TeacherMultiSelect({
  classId,
  currentTeachers,
  allUsers,
  onToggle,
  disabled,
  isOpen,
  onOpenChange,
  dropdownPos
}: {
  classId: string
  currentTeachers: { id: string; username: string }[]
  allUsers: User[]
  onToggle: (classId: string, currentTeachers: { id: string }[], teacherId: string) => void
  disabled: boolean
  isOpen: boolean
  onOpenChange: (classId: string | null, pos?: { top: number; left: number }) => void
  dropdownPos: { top: number; left: number }
}) {
  const buttonRef = useRef<HTMLButtonElement>(null)
  const currentIds = new Set(currentTeachers.map(t => t.id))

  const handleOpen = () => {
    if (!isOpen && buttonRef.current) {
      const rect = buttonRef.current.getBoundingClientRect()
      onOpenChange(classId, {
        top: rect.bottom + window.scrollY,
        left: rect.left + window.scrollX
      })
    } else {
      onOpenChange(null)
    }
  }

  return (
    <div className="relative">
      <button
        ref={buttonRef}
        type="button"
        onClick={handleOpen}
        disabled={disabled}
        className="w-full min-w-[150px] p-1.5 border rounded text-sm text-left bg-white hover:bg-gray-50 disabled:opacity-50 flex justify-between items-center"
      >
        <span className="truncate">
          {currentTeachers.length === 0
            ? "No teachers"
            : currentTeachers.map(t => t.username).join(", ")}
        </span>
        <svg className="w-4 h-4 ml-2 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
        </svg>
      </button>

      {isOpen && (
        <>
          <div
            className="fixed inset-0"
            style={{ zIndex: 9998 }}
            onClick={() => onOpenChange(null)}
          />
          <div
            className="fixed w-48 bg-white border rounded-md shadow-lg max-h-48 overflow-y-auto"
            style={{
              zIndex: 9999,
              top: dropdownPos.top,
              left: dropdownPos.left
            }}
          >
            {allUsers.length === 0 ? (
              <div className="p-2 text-sm text-gray-500">No users available</div>
            ) : (
              allUsers.map(user => (
                <label
                  key={user.id}
                  className="flex items-center gap-2 px-3 py-2 hover:bg-gray-100 cursor-pointer"
                >
                  <input
                    type="checkbox"
                    checked={currentIds.has(user.id)}
                    onChange={() => onToggle(classId, currentTeachers, user.id)}
                    className="h-4 w-4 text-blue-600 rounded"
                  />
                  <span className="text-sm">{user.username}</span>
                </label>
              ))
            )}
          </div>
        </>
      )}
    </div>
  )
}

export function ClassManagement({ subjects }: ClassManagementProps) {
  const [classes, setClasses] = useState<Class[]>([])
  const [forms, setForms] = useState<string[]>([])
  const [users, setUsers] = useState<User[]>([])
  const [loading, setLoading] = useState(true)
  const [showCreateForm, setShowCreateForm] = useState(false)
  const [className, setClassName] = useState("")
  const [formName, setFormName] = useState("")
  const [selectedSubject, setSelectedSubject] = useState("")
  const [creating, setCreating] = useState(false)
  const [editingClass, setEditingClass] = useState<Class | null>(null)
  const [savingTeachers, setSavingTeachers] = useState<string | null>(null)
  const [openDropdown, setOpenDropdown] = useState<string | null>(null)
  const [dropdownPos, setDropdownPos] = useState({ top: 0, left: 0 })

  const fetchClasses = async () => {
    setLoading(true)
    const result = await getClasses()
    if (result.success && result.classes) {
      setClasses(result.classes)
    }
    setLoading(false)
  }

  const fetchForms = async () => {
    console.log('fetchForms: Calling getAllForms')
    const result = await getAllForms()
    console.log('fetchForms: Result', result)
    if (result.success && result.forms) {
      console.log('fetchForms: Setting forms', result.forms)
      setForms(result.forms)
    } else {
      console.log('fetchForms: No forms or error', result)
    }
  }

  const fetchUsers = async () => {
    const result = await getAllUsers()
    if (result.success && result.users) {
      setUsers(result.users)
    }
  }

  useEffect(() => {
    fetchClasses()
    fetchForms()
    fetchUsers()
  }, [])

  const handleTeachersChange = async (classId: string, teacherIds: string[]) => {
    setSavingTeachers(classId)
    const result = await updateClass(classId, { teacherIds })
    setSavingTeachers(null)

    if (result.success) {
      fetchClasses()
    } else {
      alert('error' in result ? result.error : "Failed to update teachers")
    }
  }

  const toggleTeacher = (classId: string, currentTeachers: { id: string }[], teacherId: string) => {
    const currentIds = currentTeachers.map(t => t.id)
    const newIds = currentIds.includes(teacherId)
      ? currentIds.filter(id => id !== teacherId)
      : [...currentIds, teacherId]
    handleTeachersChange(classId, newIds)
  }

  const handleDropdownChange = (classId: string | null, pos?: { top: number; left: number }) => {
    setOpenDropdown(classId)
    if (pos) {
      setDropdownPos(pos)
    }
  }

  const handleCreateFromForm = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!className || !selectedSubject) return

    setCreating(true)
    const result = await createClassFromForm(className, formName, selectedSubject)
    setCreating(false)

    if (result.success) {
      alert(`Successfully created class "${className}"`)
      setShowCreateForm(false)
      setClassName("")
      setFormName("")
      setSelectedSubject("")
      fetchClasses()
    } else {
      alert('error' in result ? result.error : "Failed to create class")
    }
  }

  const handleDeleteClass = async (classId: string, className: string) => {
    if (!confirm(`Are you sure you want to delete class "${className}"? This will also delete all assignments.`)) {
      return
    }

    const result = await deleteClass(classId)
    if (result.success) {
      alert("Class deleted successfully")
      fetchClasses()
    } else {
      alert('error' in result ? result.error : "Failed to delete class")
    }
  }

  return (
    <div className="space-y-6">
      <div className="bg-white shadow rounded-lg p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-xl font-semibold">Class Management</h2>
          <button
            onClick={() => setShowCreateForm(!showCreateForm)}
            className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-md font-medium"
          >
            {showCreateForm ? "Cancel" : "Create Class"}
          </button>
        </div>

        {showCreateForm && (
          <form onSubmit={handleCreateFromForm} className="mb-6 p-4 bg-gray-50 rounded-lg">
            <h3 className="text-lg font-medium mb-4">Create New Class</h3>
            <div className="grid grid-cols-3 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Class Name *
                </label>
                <input
                  type="text"
                  value={className}
                  onChange={(e) => setClassName(e.target.value)}
                  placeholder="e.g., 7A Science"
                  className="w-full p-2 border rounded"
                  required
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Subject *
                </label>
                <select
                  value={selectedSubject}
                  onChange={(e) => setSelectedSubject(e.target.value)}
                  className="w-full p-2 border rounded"
                  required
                >
                  <option value="">Select a subject...</option>
                  {subjects.map((subject) => (
                    <option key={subject.id} value={subject.id}>
                      {subject.code} - {subject.title || subject.code}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Populate from Form
                </label>
                <select
                  value={formName}
                  onChange={(e) => setFormName(e.target.value)}
                  className="w-full p-2 border rounded"
                >
                  <option value="">No pre-population</option>
                  {forms.map((form) => (
                    <option key={form} value={form}>
                      {form}
                    </option>
                  ))}
                </select>
                <p className="text-xs text-gray-500 mt-1">
                  Optionally add all pupils from a form
                </p>
              </div>
            </div>
            <button
              type="submit"
              disabled={creating}
              className="mt-4 bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-md font-medium disabled:opacity-50"
            >
              {creating ? "Creating..." : "Create Class"}
            </button>
          </form>
        )}

        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Class Name</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Year</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Subject</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Students</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Teachers</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {loading ? (
                <tr>
                  <td colSpan={6} className="px-6 py-4 text-center text-sm text-gray-500">Loading...</td>
                </tr>
              ) : classes.length === 0 ? (
                <tr>
                  <td colSpan={6} className="px-6 py-4 text-center text-sm text-gray-500">No classes found</td>
                </tr>
              ) : (
                classes.map((cls) => (
                  <tr key={cls.id}>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">{cls.name}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{cls.year || "-"}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                      {cls.Subject.code} - {cls.Subject.title || cls.Subject.code}
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{cls._count.Assignment}</td>
                    <td className="px-6 py-4 text-sm text-gray-500">
                      <TeacherMultiSelect
                        classId={cls.id}
                        currentTeachers={cls.User}
                        allUsers={users}
                        onToggle={toggleTeacher}
                        disabled={savingTeachers === cls.id}
                        isOpen={openDropdown === cls.id}
                        onOpenChange={handleDropdownChange}
                        dropdownPos={dropdownPos}
                      />
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium space-x-3">
                      <button
                        onClick={() => setEditingClass(cls)}
                        className="text-blue-600 hover:text-blue-900"
                      >
                        Edit
                      </button>
                      <button
                        onClick={() => handleDeleteClass(cls.id, cls.name)}
                        className="text-red-600 hover:text-red-900"
                      >
                        Delete
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {editingClass && (
        <EditClassSidebar
          classData={editingClass}
          subjects={subjects}
          onClose={() => setEditingClass(null)}
          onUpdated={fetchClasses}
        />
      )}
    </div>
  )
}
