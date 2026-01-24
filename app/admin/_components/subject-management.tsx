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

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {subjects.map((subject: any) => (
          <div key={subject.id} className="relative group">
            <Link 
              href={`/hod/subject/${subject.id}`}
              className="block"
            >
              <div className="bg-white shadow rounded-lg p-6 border border-transparent hover:border-indigo-500 transition-colors min-h-[160px] flex flex-col">
                <div className="mb-2">
                  <h2 className="text-xl font-semibold text-gray-900 pr-16">{subject.code}</h2>
                  {subject.title && <h3 className="text-sm font-medium text-gray-600">{subject.title}</h3>}
                </div>
                <p className="text-gray-500 text-sm mb-4 line-clamp-2 flex-grow">{subject.studiedComment || "No introduction"}</p>
                
                <div className="flex justify-between items-center text-sm text-gray-500 mt-auto pt-4 border-t">
                  <div className="flex gap-4">
                     <span>{subject._count.classes} Classes</span>
                     <span>{subject._count.commentGroups} Groups</span>
                  </div>
                </div>
              </div>
            </Link>
            <div className="absolute bottom-4 right-4 z-10 flex gap-2">
               <SubjectUserAssignment 
                 subjectId={subject.id} 
                 assignedUsers={subject.users} 
                 allUsers={users} 
               />
               <EditSubjectForm subject={subject} />
            </div>
          </div>
        ))}

        {subjects.length === 0 && (
          <div className="col-span-full text-center py-10 bg-gray-50 rounded-lg border-2 border-dashed border-gray-200">
            <p className="text-gray-500">No subjects found. Create one to get started.</p>
          </div>
        )}
      </div>
    </div>
  )
}
