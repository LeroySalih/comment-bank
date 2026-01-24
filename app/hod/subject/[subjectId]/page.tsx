
import { prisma } from "@/lib/prisma"
import { notFound } from "next/navigation"
import Link from "next/link"
import { ClassForm } from "./_components/class-form"
import { GroupForm } from "./_components/group-form"
import { EditClassForm } from "./_components/edit-class-form"
import { EditGroupForm } from "./_components/edit-group-form"
import { TeacherAssignment } from "./_components/teacher-assignment"

import { ReorderableGroupList } from "./_components/ReorderableGroupList"
import { countWords } from "@/lib/utils"

export const dynamic = 'force-dynamic'

interface Props {
  params: Promise<{
    subjectId: string
  }>
}

import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { isAdmin } from "@/lib/access-control"

export default async function SubjectPage({ params }: Props) {
  const { subjectId } = await params
  const session = await getServerSession(authOptions)
  const userIsAdmin = isAdmin(session?.user as any)
  
  // Fetch available teachers (all users with 'teacher' role)
  const teachers = await prisma.user.findMany({
    where: {
      roles: {
        some: { name: 'teacher' }
      }
    },
    select: { id: true, username: true },
    orderBy: { username: 'asc' }
  })

  const subject = await (prisma as any).subject.findUnique({
    where: { id: subjectId },
    include: {
      classes: {
        orderBy: { name: 'asc' },
        include: { 
          _count: { select: { assignments: true } },
          teachers: { select: { id: true, username: true } }
        }
      },
      commentGroups: {
        orderBy: { displayOrder: 'asc' },
        include: { 
          options: { select: { text: true } },
          _count: { select: { options: true } } 
        }
      }
    }
  }) as any;

  if (!subject) notFound()

  const totalWordsInSubject = subject.commentGroups.reduce((acc: any, group: any) => {
    return acc + group.options.reduce((sum: any, opt: any) => sum + countWords(opt.text), 0)
  }, 0)
  const avgWordsPerGroup = subject.commentGroups.length > 0 ? (totalWordsInSubject / subject.commentGroups.length).toFixed(1) : 0

  return (
    <div className="container mx-auto py-10">
      <div className="flex items-center gap-4 mb-8">
        <Link href="/hod" className="text-indigo-600 hover:text-indigo-800">
          ← Back to Dashboard
        </Link>
        <div>
          <h1 className="text-3xl font-bold">{subject.code}</h1>
          {subject.title && <h2 className="text-xl text-gray-600">{subject.title}</h2>}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        {/* Classes Section */}
        <div className="bg-white shadow rounded-lg p-6 h-fit">
          <div className="flex justify-between items-center mb-4">
            <h2 className="text-xl font-semibold">Classes</h2>
            <ClassForm subjectId={subject.id} />
          </div>
          
          <div className="space-y-4">
            {subject.classes.map((cls: any) => (
              <div key={cls.id} className="flex justify-between items-center p-3 bg-gray-50 rounded border group/item hover:bg-gray-100 transition-colors">
                <div className="flex flex-col gap-2 w-full">
                  <div className="flex justify-between items-center w-full">
                    <div className="flex items-center gap-3">
                      <Link href={`/hod/class/${cls.id}`} className="font-medium hover:underline text-lg">
                        {cls.name}
                      </Link>
                      <EditClassForm 
                        classId={cls.id} 
                        subjectId={subject.id} 
                        initialName={cls.name} 
                      />
                    </div>
                    <div className="flex items-center gap-4">
                      <TeacherAssignment 
                        classId={cls.id}
                        currentTeachers={cls.teachers}
                        availableTeachers={teachers}
                        readOnly={!userIsAdmin}
                      />
                      <span className="text-xs text-gray-400 bg-gray-100 px-2 py-1 rounded-full whitespace-nowrap">
                        {cls._count.assignments} Pupils
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            ))}
            {subject.classes.length === 0 && (
              <p className="text-gray-500 text-sm italic">No classes added yet.</p>
            )}
          </div>
        </div>

        {/* Comment Groups Section */}
        <div className="bg-white shadow rounded-lg p-6 h-fit">
          <div className="flex justify-between items-center mb-4">
            <div>
              <h2 className="text-xl font-semibold">Comment Groups</h2>
              {subject.commentGroups.length > 0 && (
                <p className="text-sm text-gray-500">{avgWordsPerGroup} avg words per group</p>
              )}
            </div>
            <GroupForm subjectId={subject.id} />
          </div>

          <ReorderableGroupList 
            subjectId={subject.id} 
            initialGroups={subject.commentGroups} 
          />
        </div>
      </div>
    </div>
  )
}
