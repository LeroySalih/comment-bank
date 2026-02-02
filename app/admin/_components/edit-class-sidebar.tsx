"use client"

import { useState, useEffect, useRef } from "react"
import { X, ChevronDown, Search } from "lucide-react"
import { updateClass, getAllUsers } from "@/lib/server-actions/admin"
import { ClassPupilManager } from "./class-pupil-manager"

interface Class {
  id: string
  name: string
  year: string | null
  subjectId: string
  Subject: {
    code: string
    title: string | null
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
  Role: Array<{ name: string }>
}

interface EditClassSidebarProps {
  classData: Class | null
  subjects: Subject[]
  onClose: () => void
  onUpdated: () => void
}

export function EditClassSidebar({ classData, subjects, onClose, onUpdated }: EditClassSidebarProps) {
  const [name, setName] = useState("")
  const [year, setYear] = useState("")
  const [subjectId, setSubjectId] = useState("")
  const [selectedTeachers, setSelectedTeachers] = useState<string[]>([])
  const [allUsers, setAllUsers] = useState<User[]>([])
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [loadingUsers, setLoadingUsers] = useState(true)
  const [teacherDropdownOpen, setTeacherDropdownOpen] = useState(false)
  const [teacherSearch, setTeacherSearch] = useState("")
  const dropdownRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (classData) {
      setName(classData.name)
      setYear(classData.year || "")
      setSubjectId(classData.subjectId)
      setSelectedTeachers(classData.User.map(u => u.id))
    }
  }, [classData])

  useEffect(() => {
    const fetchUsers = async () => {
      setLoadingUsers(true)
      const result = await getAllUsers()
      if (result.success && result.users) {
        setAllUsers(result.users)
      }
      setLoadingUsers(false)
    }
    fetchUsers()
  }, [])

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setTeacherDropdownOpen(false)
      }
    }
    document.addEventListener("mousedown", handleClickOutside)
    return () => document.removeEventListener("mousedown", handleClickOutside)
  }, [])

  const filteredUsers = allUsers.filter(user =>
    user.username.toLowerCase().includes(teacherSearch.toLowerCase())
  )

  const selectedTeacherNames = allUsers
    .filter(u => selectedTeachers.includes(u.id))
    .map(u => u.username)

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!classData) return

    setError(null)
    setSaving(true)

    const result = await updateClass(classData.id, {
      name,
      year: year || null,
      subjectId,
      teacherIds: selectedTeachers
    })

    setSaving(false)

    if (result.success) {
      onUpdated()
      onClose()
    } else {
      setError('error' in result ? result.error : "Failed to update class")
    }
  }

  const handleTeacherToggle = (userId: string) => {
    setSelectedTeachers(prev =>
      prev.includes(userId)
        ? prev.filter(id => id !== userId)
        : [...prev, userId]
    )
  }

  if (!classData) return null

  return (
    <>
      {/* Backdrop */}
      <div
        className="fixed inset-0 bg-black/30 z-40"
        onClick={onClose}
      />

      {/* Sidebar */}
      <div className="fixed inset-y-0 right-0 w-full max-w-md bg-white shadow-xl z-50 flex flex-col">
        {/* Header */}
        <div className="flex items-center justify-between p-6 border-b">
          <h2 className="text-xl font-semibold">Edit Class</h2>
          <button
            onClick={onClose}
            className="p-2 hover:bg-gray-100 rounded-full transition-colors"
          >
            <X size={20} />
          </button>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit} className="flex-1 overflow-y-auto p-6 space-y-6">
          {/* Class Name */}
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Class Name
            </label>
            <input
              type="text"
              value={name}
              onChange={(e) => setName(e.target.value)}
              className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
              required
            />
          </div>

          {/* Year */}
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Year
            </label>
            <input
              type="text"
              value={year}
              onChange={(e) => setYear(e.target.value)}
              placeholder="e.g., 7, 8, 9"
              className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
            />
          </div>

          {/* Subject */}
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Subject
            </label>
            <select
              value={subjectId}
              onChange={(e) => setSubjectId(e.target.value)}
              className="w-full p-2 border rounded-md focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
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

          {/* Teachers */}
          <div ref={dropdownRef}>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Assigned Teachers
            </label>
            {loadingUsers ? (
              <p className="text-sm text-gray-500">Loading users...</p>
            ) : (
              <div className="relative">
                {/* Selected teachers display / dropdown trigger */}
                <button
                  type="button"
                  onClick={() => setTeacherDropdownOpen(!teacherDropdownOpen)}
                  className="w-full p-2 border rounded-md text-left flex items-center justify-between bg-white focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                >
                  <span className={selectedTeacherNames.length > 0 ? "text-gray-900" : "text-gray-500"}>
                    {selectedTeacherNames.length > 0
                      ? selectedTeacherNames.join(", ")
                      : "Select teachers..."}
                  </span>
                  <ChevronDown size={18} className={`text-gray-400 transition-transform ${teacherDropdownOpen ? "rotate-180" : ""}`} />
                </button>

                {/* Dropdown */}
                {teacherDropdownOpen && (
                  <div className="absolute z-10 w-full mt-1 bg-white border rounded-md shadow-lg">
                    {/* Search input */}
                    <div className="p-2 border-b">
                      <div className="relative">
                        <Search size={16} className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" />
                        <input
                          type="text"
                          value={teacherSearch}
                          onChange={(e) => setTeacherSearch(e.target.value)}
                          placeholder="Search teachers..."
                          className="w-full pl-8 pr-2 py-1.5 border rounded text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                          autoFocus
                        />
                      </div>
                    </div>

                    {/* Options list */}
                    <div className="max-h-48 overflow-y-auto">
                      {filteredUsers.length === 0 ? (
                        <p className="p-3 text-sm text-gray-500 text-center">No teachers found</p>
                      ) : (
                        filteredUsers.map((user) => (
                          <button
                            key={user.id}
                            type="button"
                            onClick={() => handleTeacherToggle(user.id)}
                            className={`w-full flex items-center gap-3 p-3 hover:bg-gray-50 text-left border-b last:border-b-0 ${
                              selectedTeachers.includes(user.id) ? "bg-blue-50" : ""
                            }`}
                          >
                            <input
                              type="checkbox"
                              checked={selectedTeachers.includes(user.id)}
                              onChange={() => {}}
                              className="h-4 w-4 text-blue-600 rounded focus:ring-blue-500 pointer-events-none"
                            />
                            <div className="flex-1">
                              <span className="text-sm font-medium">{user.username}</span>
                              <span className="text-xs text-gray-500 ml-2">
                                ({user.Role.map(r => r.name).join(", ")})
                              </span>
                            </div>
                          </button>
                        ))
                      )}
                    </div>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Pupil Management */}
          <ClassPupilManager classId={classData.id} />

          {error && (
            <div className="p-3 bg-red-50 border border-red-200 rounded-md">
              <p className="text-sm text-red-600">{error}</p>
            </div>
          )}
        </form>

        {/* Footer */}
        <div className="p-6 border-t bg-gray-50 flex gap-3">
          <button
            type="button"
            onClick={onClose}
            className="flex-1 px-4 py-2 border rounded-md hover:bg-gray-100 transition-colors"
          >
            Cancel
          </button>
          <button
            type="submit"
            onClick={handleSubmit}
            disabled={saving}
            className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors disabled:opacity-50"
          >
            {saving ? "Saving..." : "Save Changes"}
          </button>
        </div>
      </div>
    </>
  )
}
