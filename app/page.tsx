import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import SignOutButton from '@/components/SignOutButton';
import { getServerSession } from 'next-auth';
import { authOptions } from './api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';

export default async function Home() {
  const session = await getServerSession(authOptions);
  
  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userIsTeacher = isTeacher(session?.user);
  const userId = session?.user?.id;

  // Filter logic:
  // Admin and HOD see all subjects.
  // Teachers only see subjects with classes they are assigned to.
  const subjects = await prisma.subject.findMany({
    where: (userIsAdmin || userIsHoD) ? {} : {
       classes: {
          some: {
             teachers: {
                some: { id: userId }
             }
          }
       }
    },
    include: {
      classes: {
        where: (userIsAdmin || userIsHoD) ? {} : {
          teachers: {
            some: { id: userId }
          }
        },
        orderBy: { name: 'asc' },
        include: {
          assignments: {
            select: {
              codes: {
                select: { id: true }
              }
            }
          }
        }
      }
    },
    orderBy: { name: 'asc' }
  });

  return (
    <main className="min-h-screen p-8 bg-gray-50">
      <div className="max-w-4xl mx-auto">
        <div className="grid gap-6">
          {subjects.map((subject) => (
            <div key={subject.id} className="bg-white p-6 rounded-lg shadow-sm border border-gray-200">
              <h2 className="text-xl font-semibold mb-4 text-gray-800 flex items-center gap-2">
                <span className="bg-blue-100 text-blue-800 text-sm font-medium px-2.5 py-0.5 rounded">Subject</span>
                {subject.name}
              </h2>
              
              {subject.classes.length > 0 ? (
                <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                  {subject.classes.map((cls) => {
                    const totalAssignments = cls.assignments.length;
                    const startedAssignments = cls.assignments.filter(a => a.codes.length > 0).length;
                    const percent = totalAssignments > 0 ? Math.round((startedAssignments / totalAssignments) * 100) : 0;
                    
                    return (
                    <Link 
                      key={cls.id} 
                      href={`/class/${cls.id}`}
                      className="block p-4 border border-gray-200 rounded-md hover:border-blue-500 hover:ring-1 hover:ring-blue-500 transition-colors bg-gray-50 hover:bg-white text-center relative overflow-hidden"
                    >
                      <div className="absolute bottom-0 left-0 h-1 bg-blue-100 w-full">
                         <div className="h-full bg-blue-500 transition-all duration-500" style={{ width: `${percent}%` }} />
                      </div>
                      <div className="text-lg font-medium text-gray-900">{cls.name}</div>
                      <div className="text-sm text-gray-500">{totalAssignments} Pupils</div>
                      <div className="text-xs text-blue-600 font-semibold mt-1">{percent}% Complete</div>
                    </Link>
                  )})}
                </div>
              ) : (
                <p className="text-gray-400 italic">No classes found.</p>
              )}
            </div>
          ))}
        </div>
      </div>
    </main>
  );
}
