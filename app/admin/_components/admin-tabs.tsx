"use client"

import { useState } from "react"
import { UserManagement } from "./user-management"
import { PupilManagement } from "./pupil-management"

interface AdminTabsProps {
  users: any[]
  roles: any[]
}

export function AdminTabs({ users, roles }: AdminTabsProps) {
  const [activeTab, setActiveTab] = useState<'users' | 'pupils'>('users')

  return (
    <div className="space-y-6">
      <div className="flex border-b">
        <button
          className={`px-4 py-2 font-medium text-sm transition-colors ${
            activeTab === 'users'
              ? "border-b-2 border-blue-500 text-blue-600"
              : "text-gray-500 hover:text-gray-700"
          }`}
          onClick={() => setActiveTab('users')}
        >
          Users
        </button>
        <button
          className={`px-4 py-2 font-medium text-sm transition-colors ${
            activeTab === 'pupils'
              ? "border-b-2 border-blue-500 text-blue-600"
              : "text-gray-500 hover:text-gray-700"
          }`}
          onClick={() => setActiveTab('pupils')}
        >
          Pupils
        </button>
      </div>

      <div className="mt-6">
        {activeTab === 'users' ? (
          <UserManagement users={users} roles={roles} />
        ) : (
          <PupilManagement />
        )}
      </div>
    </div>
  )
}
