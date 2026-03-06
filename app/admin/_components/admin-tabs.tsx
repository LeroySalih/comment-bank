"use client"

import { useState } from "react"
import { UserManagement } from "./user-management"
import { PupilManagement } from "./pupil-management"
import { SubjectManagement } from "./subject-management"
import { ClassManagement } from "./class-management"
import { DeadlineManagement } from "./deadline-management"
import { ActivityLog } from "./activity-log"
import { LinkedDataUpload } from "./linked-data-upload"
import { CcgManagement } from "./ccg-management"
import { FormatManagement } from "./format-management"

interface AdminTabsProps {
  users: any[]
  roles: any[]
  subjects: any[]
  deadlines: any[]
  classes?: { id: string; name: string }[]
  commonGroups?: any[]
  commentFormatTemplate?: string
}

export function AdminTabs({ users, roles, subjects, deadlines, classes = [], commonGroups = [], commentFormatTemplate = '' }: AdminTabsProps) {
  const [activeTab, setActiveTab] = useState<'users' | 'pupils' | 'subjects' | 'classes' | 'deadlines' | 'linked-data' | 'ccg' | 'format' | 'activity'>('users')

  const tabs = [
    { key: 'users' as const, label: 'Users' },
    { key: 'pupils' as const, label: 'Pupils' },
    { key: 'classes' as const, label: 'Classes' },
    { key: 'ccg' as const, label: 'CCG' },
    { key: 'subjects' as const, label: 'Subjects' },
    { key: 'format' as const, label: 'Format' },
    { key: 'linked-data' as const, label: 'Linked Data' },
    { key: 'deadlines' as const, label: 'Deadlines' },
    { key: 'activity' as const, label: 'Activity Log' },
  ]

  return (
    <div className="space-y-6">
      <div className="flex border-b overflow-x-auto">
        {tabs.map(tab => (
          <button
            key={tab.key}
            className={`px-4 py-2 font-medium text-sm transition-colors whitespace-nowrap ${
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
        ) : activeTab === 'ccg' ? (
          <CcgManagement initialGroups={commonGroups} />
        ) : activeTab === 'format' ? (
          <FormatManagement
            initialTemplate={commentFormatTemplate}
            commonGroups={commonGroups}
            subjects={subjects}
          />
        ) : (
          <ActivityLog />
        )}
      </div>
    </div>
  )
}
