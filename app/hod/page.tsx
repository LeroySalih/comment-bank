
import { prisma } from "@/lib/prisma"
import Link from "next/link"
import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { redirect } from "next/navigation"

export const dynamic = 'force-dynamic'

export default async function HoDDashboard() {
  const session = await getServerSession(authOptions)
  if (!session?.user) {
    redirect("/login")
  }

  // Admin sees all subjects? Or just assigned ones?
  // Ideally Admin goes to Admin panel. HoD dashboard is for their assigned work.
  // But Admin might want quick access here too.
  // Let's show ALL for Admin, and ASSIGNED for others.
  
  const allowAll = session.user.roles?.some((r: any) => r.name === 'admin');
  
  const subjects = await prisma.subject.findMany({
    where: allowAll ? {} : {
      User: { some: { id: session.user.id } }
    },
    orderBy: { code: 'asc' },
    include: {
      _count: {
        select: { Class: true, CommentGroup: true }
      }
    }
  })

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark min-h-[calc(100vh-64px)]">
      {/* Page Heading */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-8 border-b border-[#f0f2f4] dark:border-gray-800 bg-white dark:bg-gray-900">
        <div className="flex min-w-72 flex-col gap-1">
          <h1 className="text-[#111418] dark:text-white text-3xl font-black leading-tight tracking-[-0.033em]">Head of Department Dashboard</h1>
          <p className="text-[#617289] dark:text-gray-400 text-base font-normal leading-normal">Manage subjects, classes, and comment banks.</p>
        </div>
        <div className="flex gap-3">
            {allowAll && (
                <button className="flex items-center justify-center rounded-xl h-10 px-4 bg-primary text-white text-sm font-bold shadow-lg shadow-primary/20 hover:bg-primary/90 transition-all">
                    <span className="material-symbols-outlined mr-2 text-lg">add</span>
                    New Subject
                </button>
            )}
        </div>
      </div>
      
      <div className="p-8">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
          {subjects.map((subject: any) => (
             <Link 
               key={subject.id}
               href={`/hod/subject/${subject.id}`}
               className="group flex flex-col bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm hover:shadow-md transition-all overflow-hidden"
             >
                <div className="p-6 flex-1 flex flex-col">
                    <div className="flex justify-between items-start mb-4">
                        <div className="size-10 rounded-lg bg-blue-50 dark:bg-blue-900/20 flex items-center justify-center text-primary">
                            <span className="material-symbols-outlined text-2xl">folder</span>
                        </div>
                        <span className="material-symbols-outlined text-gray-300 dark:text-gray-600 group-hover:text-primary transition-colors">arrow_forward</span>
                    </div>
                    
                    <h2 className="text-xl font-bold text-[#111418] dark:text-white mb-1 group-hover:text-primary transition-colors">{subject.code}</h2>
                    <h3 className="text-sm font-medium text-[#617289] dark:text-gray-400 mb-4">{subject.title}</h3>
                    
                    <p className="text-sm text-gray-500 line-clamp-2 mb-6 flex-grow">{subject.studiedComment || "No introduction set."}</p>
                    
                    <div className="flex gap-4 pt-4 border-t border-[#f0f2f4] dark:border-gray-800">
                        <div className="flex flex-col">
                            <span className="text-xs font-bold text-[#617289] dark:text-gray-500 uppercase tracking-wider">Classes</span>
                            <span className="text-lg font-bold text-[#111418] dark:text-white">{subject._count.Class}</span>
                        </div>
                        <div className="flex flex-col">
                            <span className="text-xs font-bold text-[#617289] dark:text-gray-500 uppercase tracking-wider">Groups</span>
                            <span className="text-lg font-bold text-[#111418] dark:text-white">{subject._count.CommentGroup}</span>
                        </div>
                    </div>
                </div>
             </Link>
          ))}

          {subjects.length === 0 && (
            <div className="col-span-full py-16 flex flex-col items-center justify-center text-center bg-white dark:bg-gray-900 rounded-xl border-2 border-dashed border-[#f0f2f4] dark:border-gray-800">
              <div className="size-16 bg-gray-50 dark:bg-gray-800 rounded-full flex items-center justify-center mb-4">
                  <span className="material-symbols-outlined text-gray-400 text-3xl">sentiment_dissatisfied</span>
              </div>
              <h3 className="text-lg font-bold text-[#111418] dark:text-white mb-1">No subjects found</h3>
              <p className="text-[#617289] dark:text-gray-400 max-w-sm mx-auto">You haven't been assigned to any subjects yet.</p>
            </div>
          )}
        </div>
      </div>
    </main>
  )
}
