import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import QuickGroupSelector from '@/components/QuickGroupSelector';
import CopyCommentButton from '@/components/CopyCommentButton';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';

export default async function ClassPage({ params }: { params: Promise<{ classId: string }> }) {
  const { classId } = await params;
  const session = await getServerSession(authOptions);

  const cls = await (prisma as any).class.findUnique({
    where: { id: classId },
    include: {
      teachers: {
        select: { id: true }
      },
      subject: {
        include: {
            commentGroups: {
                orderBy: { displayOrder: 'asc' },
                include: { options: true }
            }
        }
      },
      assignments: {
        where: {
          pupil: { isActive: true }
        },
        include: {
          pupil: true,
          codes: true
        }
      }
    }
  });

  if (!cls) {
    return <div>Class not found</div>;
  }

  // Authorization check
  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userIsTeacher = isTeacher(session?.user);
  
  if (userIsTeacher && !userIsAdmin && !userIsHoD) {
    const isAssigned = cls.teachers.some((t: any) => t.id === session?.user?.id);
    if (!isAssigned) {
      return <div>Class not found or access denied</div>;
    }
  }

  const groups = cls.subject.commentGroups;

  return (
    <main className="min-h-screen p-8 bg-gray-50">
      <div className="max-w-6xl mx-auto">
        <header className="mb-8">
           <Link href="/" className="inline-flex items-center text-sm text-gray-500 hover:text-gray-900 mb-4 transition-colors">
            <ArrowLeft className="w-4 h-4 mr-1" />
            Back to Dashboard
           </Link>
          <div className="flex justify-between items-end">
             <div>
                <h1 className="text-3xl font-bold text-gray-900">{cls.name}</h1>
                <p className="text-gray-600">{cls.subject.name}</p>
             </div>
             <div className="text-gray-500">
                {cls.assignments.length} Pupils
             </div>
          </div>
        </header>

        <div className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-auto max-h-[calc(100vh-250px)] relative">
             <table className="min-w-full divide-y divide-gray-200 border-separate border-spacing-0">
                <thead className="bg-gray-50">
                  <tr>
                    <th scope="col" className="sticky top-0 left-0 z-30 px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-50 border-b border-gray-200 w-[200px] min-w-[200px]">Name</th>
                    <th scope="col" className="sticky top-0 left-[200px] z-30 px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider bg-gray-50 border-b border-gray-200 w-[100px] min-w-[100px]">Gender</th>
                    {groups.map((g: any) => (
                        <th key={g.id} scope="col" className="sticky top-0 z-20 px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-24 bg-gray-50 border-b border-gray-200">
                            {g.name}
                        </th>
                    ))}
                    <th scope="col" className="sticky top-0 right-0 z-20 bg-gray-50 border-b border-gray-200 px-6 py-3">
                      <span className="sr-only">Edit</span>
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {cls.assignments.sort((a: any, b: any) => a.pupil.lastName.localeCompare(b.pupil.lastName)).map((assignment: any) => (
                    <tr key={assignment.id} className="hover:bg-gray-50 group">
                      <td className="sticky left-0 z-10 px-6 py-4 whitespace-nowrap bg-white group-hover:bg-gray-50 border-b border-gray-100 w-[200px] min-w-[200px]">
                        <div className="text-sm font-medium text-gray-900">{assignment.pupil.lastName}, {assignment.pupil.firstName}</div>
                      </td>
                      <td className="sticky left-[200px] z-10 px-6 py-4 whitespace-nowrap text-sm text-gray-500 bg-white group-hover:bg-gray-50 border-b border-gray-100 w-[100px] min-w-[100px]">
                        {assignment.pupil.gender}
                      </td>
                      {groups.map((g: any) => {
                          const currentCodeObj = assignment.codes.find((c: any) => c.groupId === g.id);
                          const currentCode = currentCodeObj?.code || null;

                          return (
                              <td key={g.id} className="px-4 py-4 whitespace-nowrap">
                                  <QuickGroupSelector 
                                    assignmentId={assignment.id} 
                                    groupId={g.id} 
                                    currentCode={currentCode}
                                    options={g.options} 
                                  />
                              </td>
                          );
                      })}
                      <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium border-b border-gray-100">
                        <div className="flex items-center justify-end gap-3">
                          <CopyCommentButton 
                            assignment={assignment} 
                            subject={cls.subject} 
                            groups={groups} 
                          />
                          <Link href={`/student/${assignment.id}`} className="text-blue-600 hover:text-blue-900 hover:underline">
                            Write Comment
                          </Link>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
             </table>
        </div>
      </div>
    </main>
  );
}
