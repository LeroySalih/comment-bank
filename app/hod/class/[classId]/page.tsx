
import { prisma } from "@/lib/prisma"
import { notFound } from "next/navigation"
import Link from "next/link"
import { PupilForm } from "./_components/pupil-form"
import { StudentRow } from "./_components/student-row"

export const dynamic = 'force-dynamic'

interface Props {
  params: Promise<{
    classId: string
  }>
}

export default async function ClassPage({ params }: Props) {
  const { classId } = await params
  
  const classData = await prisma.class.findUnique({
    where: { id: classId },
    include: {
      students: {
        orderBy: { lastName: 'asc' }
      },
      course: true
    }
  })

  if (!classData) notFound()

  return (
    <div className="container mx-auto py-10">
      <div className="flex flex-col gap-2 mb-8">
        <div className="flex items-center gap-2 text-sm text-gray-500">
           <Link href="/hod" className="hover:text-indigo-600">Dashboard</Link>
           <span>/</span>
           <Link href={`/hod/subject/${classData.courseId}`} className="hover:text-indigo-600">{classData.course.name}</Link>
           <span>/</span>
           <span>{classData.name}</span>
        </div>
        <h1 className="text-3xl font-bold">{classData.name} - Pupils</h1>
      </div>

      <div className="bg-white shadow rounded-lg p-6">
        <div className="flex justify-between items-center mb-6">
           <h2 className="text-xl font-semibold">Pupils</h2>
           <PupilForm classId={classData.id} />
        </div>

        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Gender</th>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Target Level</th>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Actual Level</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {classData.students.map((student) => (
                <StudentRow 
                  key={student.id} 
                  student={student} 
                  classId={classData.id} 
                />
              ))}
              {classData.students.length === 0 && (
                <tr>
                  <td colSpan={3} className="px-6 py-10 text-center text-sm text-gray-500">
                    No pupils added to this class yet.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
