"use client"

import { useState } from "react"
import { updateUserRoles, updateUserActiveStatus } from "@/lib/server-actions/admin"

type User = {
  id: string
  username: string
  isActive: boolean
  Role: { id: string; name: string }[]
}

type Role = {
  id: string
  name: string
}

export function UserTable({ users, availableRoles }: { users: User[]; availableRoles: Role[] }) {
  const [loadingId, setLoadingId] = useState<string | null>(null)
  const [filterUsername, setFilterUsername] = useState("")

  const filteredUsers = users.filter(u =>
    u.username.toLowerCase().includes(filterUsername.toLowerCase())
  )

  const handleRoleChange = async (userId: string, roleName: string, isChecked: boolean) => {
    setLoadingId(userId)
    const user = users.find(u => u.id === userId)
    if (!user) return

    const currentRoleNames = user.Role.map(r => r.name)
    let newRoleNames: string[]

    if (isChecked) {
      newRoleNames = [...currentRoleNames, roleName]
    } else {
      newRoleNames = currentRoleNames.filter(r => r !== roleName)
    }

    await updateUserRoles(userId, newRoleNames)
    setLoadingId(null)
  }

  const handleActiveStatusChange = async (userId: string, isActive: boolean) => {
    setLoadingId(userId)
    await updateUserActiveStatus(userId, isActive)
    setLoadingId(null)
  }

  return (
    <div>
      <div className="flex gap-3 mb-4">
        <input
          type="text"
          placeholder="Filter by username…"
          value={filterUsername}
          onChange={e => setFilterUsername(e.target.value)}
          className="flex-1 max-w-xs px-3 py-2 border rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
        {filterUsername && (
          <button
            onClick={() => setFilterUsername("")}
            className="px-3 py-2 text-sm text-gray-500 hover:text-gray-800 border rounded-md"
          >
            Clear
          </button>
        )}
      </div>
      <p className="text-xs text-gray-400 mb-3">Showing {filteredUsers.length} of {users.length} users</p>
    <div className="overflow-x-auto">
      <table className="min-w-full divide-y divide-gray-200">
        <thead className="bg-gray-50">
          <tr>
            <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Username</th>
            <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
            <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Roles</th>
            <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Manage Roles</th>
          </tr>
        </thead>
        <tbody className="bg-white divide-y divide-gray-200">
          {filteredUsers.map((user) => (
            <tr key={user.id} className={!user.isActive ? "bg-gray-50" : ""}>
              <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                {user.username}
              </td>
              <td className="px-6 py-4 whitespace-nowrap text-sm">
                <button
                  onClick={() => handleActiveStatusChange(user.id, !user.isActive)}
                  disabled={loadingId === user.id}
                  className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${
                    user.isActive
                      ? "bg-green-100 text-green-800 hover:bg-green-200"
                      : "bg-red-100 text-red-800 hover:bg-red-200"
                  } ${loadingId === user.id ? "opacity-50 cursor-not-allowed" : "cursor-pointer"}`}
                >
                  {user.isActive ? "Active" : "Inactive"}
                </button>
              </td>
              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                {user.Role.map(r => r.name).join(", ") || "No roles"}
              </td>
              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                <div className="flex gap-4">
                  {availableRoles.map((role) => (
                    <label key={role.id} className="inline-flex items-center">
                      <input
                        type="checkbox"
                        disabled={loadingId === user.id}
                        checked={user.Role.some(r => r.name === role.name)}
                        onChange={(e) => handleRoleChange(user.id, role.name, e.target.checked)}
                        className="rounded border-gray-300 text-indigo-600 shadow-sm focus:border-indigo-300 focus:ring focus:ring-indigo-200 focus:ring-opacity-50"
                      />
                      <span className="ml-2 capitalize">{role.name}</span>
                    </label>
                  ))}
                </div>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
    </div>
  )
}
