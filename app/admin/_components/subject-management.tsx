"use client"

import { SubjectForm } from "./subjects/subject-form"
import { EditSubjectForm } from "./subjects/edit-subject-form"
import { SubjectUserAssignment } from "./subjects/subject-user-assignment"
import Link from "next/link"

interface SubjectManagementProps {
  subjects: any[]
  users: any[]
}

export function SubjectManagement({ subjects, users }: SubjectManagementProps) {
  return (
    <div>
      <div className="mb-8">
        <SubjectForm />
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-gray-200 dark:border-gray-700 text-left">
              <th className="px-4 py-3 font-semibold text-gray-600 dark:text-gray-400">Code</th>
              <th className="px-4 py-3 font-semibold text-gray-600 dark:text-gray-400">Title</th>
              <th className="px-4 py-3 font-semibold text-gray-600 dark:text-gray-400 text-center">Classes</th>
              <th className="px-4 py-3 font-semibold text-gray-600 dark:text-gray-400 text-center">Groups</th>
              <th className="px-4 py-3 font-semibold text-gray-600 dark:text-gray-400">Assigned Users</th>
              <th className="px-4 py-3 font-semibold text-gray-600 dark:text-gray-400 text-right">Actions</th>
            </tr>
          </thead>
          <tbody>
            {subjects.map((subject: any) => (
              <tr key={subject.id} className="border-b border-gray-100 dark:border-gray-800 hover:bg-gray-50 dark:hover:bg-gray-800/50 transition-colors">
                <td className="px-4 py-3">
                  <Link href={`/hod/subject/${subject.id}`} className="font-semibold text-indigo-600 dark:text-indigo-400 hover:underline">
                    {subject.code}
                  </Link>
                </td>
                <td className="px-4 py-3 text-gray-700 dark:text-gray-300">{subject.title || "—"}</td>
                <td className="px-4 py-3 text-center text-gray-600 dark:text-gray-400">{subject._count.Class}</td>
                <td className="px-4 py-3 text-center text-gray-600 dark:text-gray-400">{subject._count.CommentGroup}</td>
                <td className="px-4 py-3">
                  <div className="flex flex-wrap gap-1">
                    {subject.User?.length > 0 ? subject.User.map((u: any) => (
                      <span key={u.id} className="text-xs bg-gray-100 dark:bg-gray-800 text-gray-600 dark:text-gray-400 px-1.5 py-0.5 rounded">
                        {u.username}
                      </span>
                    )) : <span className="text-xs text-gray-400 italic">None</span>}
                  </div>
                </td>
                <td className="px-4 py-3 text-right">
                  <div className="flex gap-2 justify-end">
                    <SubjectUserAssignment
                      subjectId={subject.id}
                      assignedUsers={subject.User}
                      allUsers={users}
                    />
                    <EditSubjectForm subject={subject} />
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {subjects.length === 0 && (
          <div className="text-center py-10 bg-gray-50 dark:bg-gray-800 rounded-lg border-2 border-dashed border-gray-200 dark:border-gray-700 mt-4">
            <p className="text-gray-500">No subjects found. Create one to get started.</p>
          </div>
        )}
      </div>
    </div>
  )
}
