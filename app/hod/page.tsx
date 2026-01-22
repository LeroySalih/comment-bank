
import { prisma } from "@/lib/prisma"
import { SubjectForm } from "./_components/subject-form"
import { EditSubjectForm } from "./_components/edit-subject-form"
import Link from "next/link"

export const dynamic = 'force-dynamic'

export default async function HoDDashboard() {
  const subjects = await prisma.course.findMany({
    orderBy: { name: 'asc' },
    include: {
      _count: {
        select: { classes: true, commentGroups: true }
      }
    }
  })

  return (
    <div className="container mx-auto py-10">
      <h1 className="text-3xl font-bold mb-8">Head of Department Dashboard</h1>
      
      <div className="mb-8">
        <SubjectForm />
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {subjects.map(subject => (
          <div key={subject.id} className="relative group">
            <Link 
              href={`/hod/subject/${subject.id}`}
              className="block"
            >
              <div className="bg-white shadow rounded-lg p-6 border border-transparent hover:border-indigo-500 transition-colors min-h-[160px] flex flex-col">
                <h2 className="text-xl font-semibold mb-2 text-gray-900 pr-16">{subject.name}</h2>
                <p className="text-gray-500 text-sm mb-4 line-clamp-2 flex-grow">{subject.studiedComment || "No introduction"}</p>
                
                <div className="flex justify-between text-sm text-gray-500 mt-auto pt-4 border-t">
                  <span>{subject._count.classes} Classes</span>
                  <span>{subject._count.commentGroups} Groups</span>
                </div>
              </div>
            </Link>
            <EditSubjectForm subject={subject} />
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
