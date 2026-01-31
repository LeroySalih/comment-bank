
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
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark min-h-[calc(100vh-64px)]">
      {/* Header */}
      <div className="bg-white dark:bg-gray-900 border-b border-[#f0f2f4] dark:border-gray-800 px-8 py-6">
          <Link href="/hod" className="inline-flex items-center text-sm font-medium text-[#617289] hover:text-primary transition-colors mb-4">
              <span className="material-symbols-outlined text-lg mr-1">arrow_back</span>
              Back to Department
          </Link>
          <div className="flex justify-between items-end">
             <div>
                <h1 className="text-3xl font-black text-[#111418] dark:text-white leading-tight tracking-tight">{subject.code}</h1>
                <p className="text-[#617289] dark:text-gray-400 mt-1 text-lg">{subject.title}</p>
             </div>
             <div className="flex gap-4">
                 <div className="text-right hidden md:block">
                     <p className="text-sm text-[#617289] uppercase font-bold tracking-wider">Avg. Words</p>
                     <p className="text-xl font-black text-[#111418] dark:text-white">{avgWordsPerGroup}</p>
                 </div>
                 <div className="h-10 w-px bg-gray-200 dark:bg-gray-700 hidden md:block"></div>
                 <div className="text-right hidden md:block">
                     <p className="text-sm text-[#617289] uppercase font-bold tracking-wider">Groups</p>
                     <p className="text-xl font-black text-[#111418] dark:text-white">{subject.commentGroups.length}</p>
                 </div>
             </div>
          </div>
      </div>

      <div className="p-8 grid grid-cols-1 xl:grid-cols-2 gap-8">
        {/* Classes Section */}
        <section className="bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm flex flex-col overflow-hidden h-fit">
            <div className="p-6 border-b border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center bg-gray-50/50 dark:bg-gray-800/50">
                <div className="flex items-center gap-3">
                    <div className="size-8 rounded-lg bg-blue-100 dark:bg-blue-900/30 flex items-center justify-center text-primary">
                        <span className="material-symbols-outlined text-lg">school</span>
                    </div>
                    <h2 className="text-lg font-bold text-[#111418] dark:text-white">Classes</h2>
                </div>
                <ClassForm subjectId={subject.id} />
            </div>
            
            <div className="p-2 overflow-y-auto max-h-[600px]">
                <div className="space-y-1">
                    {subject.classes.map((cls: any) => (
                      <div key={cls.id} className="group p-4 rounded-lg hover:bg-gray-50 dark:hover:bg-gray-800 transition-colors border border-transparent hover:border-[#f0f2f4] dark:hover:border-gray-700">
                        <div className="flex justify-between items-start">
                           <div className="flex-1">
                               <div className="flex items-center gap-2 mb-1">
                                    <Link href={`/hod/class/${cls.id}`} className="text-base font-bold text-[#111418] dark:text-white hover:text-primary transition-colors flex items-center gap-1">
                                        {cls.name}
                                        <span className="material-symbols-outlined text-gray-300 text-sm opacity-0 group-hover:opacity-100 transition-opacity">open_in_new</span>
                                    </Link>
                                    <EditClassForm 
                                        classId={cls.id} 
                                        subjectId={subject.id} 
                                        initialName={cls.name} 
                                      />
                               </div>
                               <div className="flex items-center gap-2 mt-2">
                                   <TeacherAssignment 
                                        classId={cls.id}
                                        currentTeachers={cls.teachers}
                                        availableTeachers={teachers}
                                        readOnly={!userIsAdmin}
                                      />
                               </div>
                           </div>
                           <span className="bg-gray-100 dark:bg-gray-800 text-gray-600 dark:text-gray-300 text-xs font-bold px-2 py-1 rounded-md">
                                {cls._count.assignments} Pupils
                           </span>
                        </div>
                      </div>
                    ))}
                    {subject.classes.length === 0 && (
                        <div className="py-12 text-center text-gray-400 italic">No classes found. Add one to get started.</div>
                    )}
                </div>
            </div>
        </section>

        {/* Comment Groups Section */}
        <section className="bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm flex flex-col overflow-hidden h-fit">
           <div className="p-6 border-b border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center bg-gray-50/50 dark:bg-gray-800/50">
                <div className="flex items-center gap-3">
                    <div className="size-8 rounded-lg bg-emerald-100 dark:bg-emerald-900/30 flex items-center justify-center text-emerald-600">
                        <span className="material-symbols-outlined text-lg">library_books</span>
                    </div>
                    <h2 className="text-lg font-bold text-[#111418] dark:text-white">Comment Bank</h2>
                </div>
                <GroupForm subjectId={subject.id} />
            </div>

            <div className="p-6">
                 <ReorderableGroupList 
                    subjectId={subject.id} 
                    initialGroups={subject.commentGroups} 
                  />
            </div>
        </section>
      </div>
    </main>
  )
}
