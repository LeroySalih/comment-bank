"use client"

import { useState } from "react"
import { UserManagement } from "./user-management"
import { PupilManagement } from "./pupil-management"
import { SubjectManagement } from "./subject-management"
import { ClassManagement } from "./class-management"

interface AdminTabsProps {
  users: any[]
  roles: any[]
  subjects: any[]
}

export function AdminTabs({ users, roles, subjects }: AdminTabsProps) {
  const [activeTab, setActiveTab] = useState<'users' | 'pupils' | 'subjects' | 'classes'>('users')

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
        <button
          className={`px-4 py-2 font-medium text-sm transition-colors ${
            activeTab === 'subjects'
              ? "border-b-2 border-blue-500 text-blue-600"
              : "text-gray-500 hover:text-gray-700"
          }`}
          onClick={() => setActiveTab('subjects')}
        >
          Subjects
        </button>
        <button
          className={`px-4 py-2 font-medium text-sm transition-colors ${
            activeTab === 'classes'
              ? "border-b-2 border-blue-500 text-blue-600"
              : "text-gray-500 hover:text-gray-700"
          }`}
          onClick={() => setActiveTab('classes')}
        >
          Classes
        </button>
      </div>

      <div className="mt-6">
        {activeTab === 'users' ? (
          <UserManagement users={users} roles={roles} />
        ) : activeTab === 'pupils' ? (
          <PupilManagement />
        ) : activeTab === 'subjects' ? (
          <SubjectManagement subjects={subjects} users={users} />
        ) : (
          <ClassManagement subjects={subjects} />
        )}
      </div>
    </div>
  )
}
