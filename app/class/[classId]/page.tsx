import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import QuickGroupSelector from '@/components/QuickGroupSelector';
import CopyCommentButton from '@/components/CopyCommentButton';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';

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

  // Decrypt pupil names for display
  cls.assignments = cls.assignments.map((assignment: any) => ({
    ...assignment,
    pupil: {
      ...assignment.pupil,
      firstName: decrypt(assignment.pupil.firstName),
      lastName: decrypt(assignment.pupil.lastName)
    }
  }));

  // Re-sort because database sort was on encrypted strings
  cls.assignments.sort((a: any, b: any) => 
    a.pupil.lastName.localeCompare(b.pupil.lastName)
  );

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark h-[calc(100vh-64px)] overflow-hidden">
        {/* Page Heading */}
        <div className="flex flex-wrap items-center justify-between gap-4 p-6 bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748]">
            <div className="flex flex-col gap-1">
                <div className="flex items-center gap-2">
                    <Link href="/" className="text-gray-500 hover:text-gray-900 transition-colors">
                        <ArrowLeft className="w-5 h-5" />
                    </Link>
                    <h1 className="text-[#111418] dark:text-white text-3xl font-black leading-tight tracking-[-0.033em]">
                        Class Matrix: {cls.name}
                    </h1>
                </div>
                <p className="text-[#617289] dark:text-gray-400 text-sm font-normal leading-normal">
                    {cls.subject.code} • {cls.assignments.length} Students
                </p>
            </div>
            <div className="flex gap-3">
                <button className="flex items-center justify-center rounded-lg h-10 px-4 bg-background-light dark:bg-[#2d3748] text-[#111418] dark:text-white text-sm font-bold tracking-[0.015em] border border-[#dbe0e6] dark:border-[#3a4454] hover:bg-gray-100 dark:hover:bg-[#3a4454] transition-colors shadow-sm">
                    <span className="material-symbols-outlined mr-2 text-lg">download</span>
                    Export CSV
                </button>
            </div>
        </div>

        {/* Table Content Container */}
        <div className="flex-1 overflow-auto p-6">
            <div className="bg-white dark:bg-[#1a222c] rounded-xl border border-[#dbe0e6] dark:border-[#2d3748] shadow-sm overflow-hidden h-full flex flex-col">
                <div className="overflow-auto relative h-full">
                    <table className="w-full text-left border-separate border-spacing-0">
                        <thead className="bg-white dark:bg-[#1a222c] sticky top-0 z-30">
                            <tr>
                                <th scope="col" className="sticky top-0 left-0 z-40 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] w-[240px] min-w-[240px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                    Student Name
                                </th>
                                <th scope="col" className="sticky top-0 left-[240px] z-40 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] w-[100px] min-w-[100px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                    Gender
                                </th>
                                {groups.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        {g.name}
                                    </th>
                                ))}
                                <th scope="col" className="sticky top-0 right-0 z-30 px-6 py-4 text-[#617289] dark:text-gray-400 text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] text-right min-w-[120px]">
                                    Actions
                                </th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-[#e5e7eb] dark:divide-[#2d3748]">
                            {cls.assignments.sort((a: any, b: any) => a.pupil.lastName.localeCompare(b.pupil.lastName)).map((assignment: any) => (
                                <tr key={assignment.id} className="group hover:bg-primary/5 dark:hover:bg-primary/10 transition-colors">
                                    <td className="sticky left-0 z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-primary/5 dark:group-hover:bg-gray-800/50 border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                        <div className="flex flex-col">
                                            <span className="text-[#111418] dark:text-white text-sm font-semibold">{assignment.pupil.lastName}, {assignment.pupil.firstName}</span>
                                        </div>
                                    </td>
                                    <td className="sticky left-[240px] z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-primary/5 dark:group-hover:bg-gray-800/50 border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                        <span className="text-sm text-[#617289] dark:text-gray-400">{assignment.pupil.gender}</span>
                                    </td>
                                    {groups.map((g: any) => {
                                        const currentCodeObj = assignment.codes.find((c: any) => c.groupId === g.id);
                                        const currentCode = currentCodeObj?.code || null;
                                        return (
                                            <td key={g.id} className="px-6 py-4 whitespace-nowrap">
                                                <QuickGroupSelector 
                                                    assignmentId={assignment.id} 
                                                    groupId={g.id} 
                                                    currentCode={currentCode}
                                                    options={g.options} 
                                                    context={{
                                                        firstName: assignment.pupil.firstName,
                                                        gender: assignment.pupil.gender,
                                                        subjectTitle: cls.subject.title || undefined,
                                                        year: cls.year || undefined,
                                                        eoyLevel: assignment.eoyLevel,
                                                        targetLevel: assignment.targetLevel
                                                    }}
                                                />
                                            </td>
                                        );
                                    })}
                                    <td className="px-6 py-4 whitespace-nowrap text-right">
                                        <div className="flex items-center justify-end gap-3">
                                            <CopyCommentButton 
                                                assignment={assignment} 
                                                subject={cls.subject} 
                                                groups={groups} 
                                            />
                                            <Link href={`/student/${assignment.id}`} className="text-primary hover:text-blue-700 text-sm font-bold transition-colors inline-flex items-center gap-1">
                                                <span className="material-symbols-outlined text-lg">visibility</span>
                                                Preview
                                            </Link>
                                        </div>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </main>
  );
}
