"use client"

import { useState } from "react"
import { updateUserRoles } from "@/lib/server-actions/admin"

type User = {
  id: string
  username: string
  roles: { id: string; name: string }[]
}

type Role = {
  id: string
  name: string
}

export function UserTable({ users, availableRoles }: { users: User[]; availableRoles: Role[] }) {
  const [loadingId, setLoadingId] = useState<string | null>(null)

  const handleRoleChange = async (userId: string, roleName: string, isChecked: boolean) => {
    setLoadingId(userId)
    const user = users.find(u => u.id === userId)
    if (!user) return

    const currentRoleNames = user.roles.map(r => r.name)
    let newRoleNames: string[]

    if (isChecked) {
      newRoleNames = [...currentRoleNames, roleName]
    } else {
      newRoleNames = currentRoleNames.filter(r => r !== roleName)
    }

    await updateUserRoles(userId, newRoleNames)
    setLoadingId(null)
  }

  return (
    <div className="overflow-x-auto">
      <table className="min-w-full divide-y divide-gray-200">
        <thead className="bg-gray-50">
          <tr>
            <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Username</th>
            <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Roles</th>
            <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Manage</th>
          </tr>
        </thead>
        <tbody className="bg-white divide-y divide-gray-200">
          {users.map((user) => (
            <tr key={user.id}>
              <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">{user.username}</td>
              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                {user.roles.map(r => r.name).join(", ")}
              </td>
              <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                <div className="flex gap-4">
                  {availableRoles.map((role) => (
                    <label key={role.id} className="inline-flex items-center">
                      <input
                        type="checkbox"
                        disabled={loadingId === user.id}
                        checked={user.roles.some(r => r.name === role.name)}
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
  )
}
