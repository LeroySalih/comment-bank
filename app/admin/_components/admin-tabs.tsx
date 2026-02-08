"use client"

import { useState } from "react"
import { UserManagement } from "./user-management"
import { PupilManagement } from "./pupil-management"
import { SubjectManagement } from "./subject-management"
import { ClassManagement } from "./class-management"
import { DeadlineManagement } from "./deadline-management"
import { ActivityLog } from "./activity-log"
import { LinkedDataUpload } from "./linked-data-upload"

interface AdminTabsProps {
  users: any[]
  roles: any[]
  subjects: any[]
  deadlines: any[]
  classes?: { id: string; name: string }[]
}

export function AdminTabs({ users, roles, subjects, deadlines, classes = [] }: AdminTabsProps) {
  const [activeTab, setActiveTab] = useState<'users' | 'pupils' | 'subjects' | 'classes' | 'deadlines' | 'linked-data' | 'activity'>('users')

  const tabs = [
    { key: 'users' as const, label: 'Users' },
    { key: 'pupils' as const, label: 'Pupils' },
    { key: 'subjects' as const, label: 'Subjects' },
    { key: 'classes' as const, label: 'Classes' },
    { key: 'deadlines' as const, label: 'Deadlines' },
    { key: 'linked-data' as const, label: 'Linked Data' },
    { key: 'activity' as const, label: 'Activity Log' },
  ]

  return (
    <div className="space-y-6">
      <div className="flex border-b">
        {tabs.map(tab => (
          <button
            key={tab.key}
            className={`px-4 py-2 font-medium text-sm transition-colors ${
              activeTab === tab.key
                ? "border-b-2 border-blue-500 text-blue-600"
                : "text-gray-500 hover:text-gray-700"
            }`}
            onClick={() => setActiveTab(tab.key)}
          >
            {tab.label}
          </button>
        ))}
      </div>

      <div className="mt-6">
        {activeTab === 'users' ? (
          <UserManagement users={users} roles={roles} />
        ) : activeTab === 'pupils' ? (
          <PupilManagement />
        ) : activeTab === 'subjects' ? (
          <SubjectManagement subjects={subjects} users={users} />
        ) : activeTab === 'classes' ? (
          <ClassManagement subjects={subjects} />
        ) : activeTab === 'deadlines' ? (
          <DeadlineManagement deadlines={deadlines} />
        ) : activeTab === 'linked-data' ? (
          <LinkedDataUpload classes={classes} />
        ) : (
          <ActivityLog />
        )}
      </div>
    </div>
  )
}
