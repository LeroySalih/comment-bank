"use client"

import { useState, useRef, useEffect } from "react"
import { assignUserToSubject, removeUserFromSubject } from "@/lib/server-actions/admin"
import { useRouter } from "next/navigation"
import { Users, Check, ChevronDown, X, UserPlus, UserMinus } from "lucide-react"

interface User {
  id: string
  username: string
  roles: { name: string }[]
}

interface SubjectUserAssignmentProps {
  subjectId: string
  assignedUsers: User[]
  allUsers: User[]
}

export function SubjectUserAssignment({ subjectId, assignedUsers, allUsers }: SubjectUserAssignmentProps) {
  const [isOpen, setIsOpen] = useState(false)
  const [search, setSearch] = useState("")
  const [isLoading, setIsLoading] = useState(false)
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

  async function handleToggle(userId: string, isAssigned: boolean) {
    setIsLoading(true)
    let result
    if (isAssigned) {
      result = await removeUserFromSubject(subjectId, userId)
    } else {
      result = await assignUserToSubject(subjectId, userId)
    }
    setIsLoading(false)

    if (result.success) {
      router.refresh()
    } else {
      alert(result.error || "Failed to update assignment")
    }
  }

  const filteredUsers = allUsers.filter(u => 
    u.username.toLowerCase().includes(search.toLowerCase())
  )

  const usersByRole = filteredUsers.reduce((acc, user) => {
    const role = user.roles[0]?.name || "user"
    if (!acc[role]) acc[role] = []
    acc[role].push(user)
    return acc
  }, {} as Record<string, User[]>)

  return (
    <div className="relative" ref={dropdownRef} onClick={(e) => e.stopPropagation()}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        className="flex items-center gap-2 text-sm text-indigo-600 hover:text-indigo-800 transition-colors"
        title="Manage Users"
      >
        <Users size={16} />
        <span>{assignedUsers.length} Users</span>
      </button>

      {isOpen && (
        <div className="absolute top-full left-0 mt-2 z-20 w-80 bg-white border border-gray-200 rounded-lg shadow-xl overflow-hidden flex flex-col max-h-[400px]">
          <div className="p-3 border-b bg-gray-50 flex justify-between items-center sticky top-0">
            <h3 className="text-sm font-semibold text-gray-700">Manage Access</h3>
            <button onClick={() => setIsOpen(false)} className="text-gray-400 hover:text-gray-600">
              <X size={16} />
            </button>
          </div>
          
          <div className="p-2 border-b">
            <input 
              type="text" 
              placeholder="Search users..." 
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full text-sm border-gray-300 rounded-md focus:ring-indigo-500 focus:border-indigo-500 p-1.5 border"
            />
          </div>

          <div className="overflow-y-auto flex-grow p-2 space-y-4">
            {Object.entries(usersByRole).map(([role, users]) => (
              <div key={role}>
                <h4 className="text-xs font-bold text-gray-400 uppercase mb-2 px-2">{role}s</h4>
                <div className="space-y-1">
                  {users.map(user => {
                    const isAssigned = assignedUsers.some(u => u.id === user.id)
                    return (
                      <div key={user.id} className="flex items-center justify-between p-2 hover:bg-gray-50 rounded-md group">
                        <div className="flex flex-col">
                          <span className="text-sm font-medium text-gray-700">{user.username}</span>
                        </div>
                        <button
                          onClick={() => handleToggle(user.id, isAssigned)}
                          disabled={isLoading}
                          className={`p-1 rounded-full transition-colors ${
                            isAssigned 
                              ? "bg-green-100 text-green-600 hover:bg-red-100 hover:text-red-600" 
                              : "bg-gray-100 text-gray-400 hover:bg-indigo-100 hover:text-indigo-600"
                          }`}
                          title={isAssigned ? "Remove Access" : "Grant Access"}
                        >
                          {isAssigned ? <Check size={16} /> : <UserPlus size={16} />}
                        </button>
                      </div>
                    )
                  })}
                </div>
              </div>
            ))}
            {filteredUsers.length === 0 && (
              <p className="text-sm text-gray-500 text-center py-4">No users found.</p>
            )}
          </div>
        </div>
      )}
    </div>
  )
}
