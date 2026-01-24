
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
  
  const subjects = await (prisma as any).subject.findMany({
    where: allowAll ? {} : {
      users: { some: { id: session.user.id } }
    },
    orderBy: { code: 'asc' },
    include: {
      _count: {
        select: { classes: true, commentGroups: true }
      }
    }
  })

  return (
    <div className="container mx-auto py-10">
      <h1 className="text-3xl font-bold mb-8">Head of Department Dashboard</h1>
      
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
                
                <div className="flex justify-between text-sm text-gray-500 mt-auto pt-4 border-t">
                  <span>{subject._count.classes} Classes</span>
                  <span>{subject._count.commentGroups} Groups</span>
                </div>
              </div>
            </Link>
          </div>
        ))}

        {subjects.length === 0 && (
          <div className="col-span-full text-center py-10 bg-gray-50 rounded-lg border-2 border-dashed border-gray-200">
            <p className="text-gray-500">No subjects assigned to you.</p>
          </div>
        )}
      </div>
    </div>
  )
}
