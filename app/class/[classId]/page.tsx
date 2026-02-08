import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import Tooltip from '@/components/Tooltip';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';
import StudentMatrixRow from './_components/StudentMatrixRow';

export default async function ClassPage({ params }: { params: Promise<{ classId: string }> }) {
  const { classId } = await params;
  const session = await getServerSession(authOptions);

  const cls = await prisma.class.findUnique({
    where: { id: classId },
    include: {
      User: {
        select: { id: true }
      },
      Subject: {
        include: {
            CommentGroup: {
                orderBy: { displayOrder: 'asc' },
                include: { CommentOption: true }
            }
        }
      },
      Assignment: {
        where: {
          Pupil: { isActive: true }
        },
        include: {
          Pupil: true,
          PupilCode: true,
          CommonPupilCode: true
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
    const isAssigned = cls.User.some((t) => t.id === session?.user?.id);
    if (!isAssigned) {
      return <div>Class not found or access denied</div>;
    }
  }

  const groups = cls.Subject.CommentGroup;

  // Fetch Common Comment Groups
  const commonGroups = await (prisma as any).commonCommentGroup.findMany({
    orderBy: [
      { paragraphPosition: 'asc' },
      { displayOrder: 'asc' }
    ],
    include: {
      CommonCommentOption: {
        orderBy: { displayOrder: 'asc' }
      }
    }
  });

  // Fetch wrapper template
  const wrapperSetting = await (prisma as any).appSetting.findUnique({
    where: { key: 'p2_wrapper_template' }
  });
  const wrapperTemplate = wrapperSetting?.value || '';

  // Decrypt pupil names for display
  const assignments = cls.Assignment.map((assignment: any) => ({
    ...assignment,
    Pupil: {
      ...assignment.Pupil,
      firstName: decrypt(assignment.Pupil.firstName),
      lastName: decrypt(assignment.Pupil.lastName)
    },
    finalComment: assignment.finalComment ? decrypt(assignment.finalComment) : null
  }));

  // Re-sort because database sort was on encrypted strings
  assignments.sort((a: any, b: any) =>
    a.Pupil.lastName.localeCompare(b.Pupil.lastName)
  );

  // Organize CCGs by paragraph position for column headers
  const p1Groups = commonGroups.filter((g: any) => g.paragraphPosition === 'p1');
  const p2Groups = commonGroups.filter((g: any) => g.paragraphPosition === 'p2');
  const p4Groups = commonGroups.filter((g: any) => g.paragraphPosition === 'p4');

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
                    {cls.Subject.code} • {assignments.length} Students
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
                                <th scope="col" className="sticky top-0 left-[240px] z-40 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] w-[80px] min-w-[80px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                    Gender
                                </th>
                                <th scope="col" className="sticky top-0 left-[320px] z-40 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] w-[140px] min-w-[140px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                    Status
                                </th>
                                {/* P1 CCG columns */}
                                {p1Groups.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-green-700 dark:text-green-400 text-xs font-bold uppercase tracking-wider bg-green-50/50 dark:bg-green-900/10 border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        <Tooltip content={`P1 — ${g.name}`}>
                                            {g.name}
                                        </Tooltip>
                                    </th>
                                ))}
                                {/* P2 CCG columns */}
                                {p2Groups.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-green-700 dark:text-green-400 text-xs font-bold uppercase tracking-wider bg-green-50/50 dark:bg-green-900/10 border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        <Tooltip content={`P2 — ${g.name}`}>
                                            {g.name}
                                        </Tooltip>
                                    </th>
                                ))}
                                {/* Subject-specific group columns */}
                                {groups.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        <Tooltip content={g.name}>
                                            {g.name}
                                        </Tooltip>
                                    </th>
                                ))}
                                {/* P4 CCG columns */}
                                {p4Groups.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-green-700 dark:text-green-400 text-xs font-bold uppercase tracking-wider bg-green-50/50 dark:bg-green-900/10 border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        <Tooltip content={`P4 — ${g.name}`}>
                                            {g.name}
                                        </Tooltip>
                                    </th>
                                ))}
                                <th scope="col" className="sticky top-0 right-0 z-30 px-6 py-4 text-[#617289] dark:text-gray-400 text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] text-right min-w-[120px]">
                                    Actions
                                </th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-[#e5e7eb] dark:divide-[#2d3748]">
                            {assignments.map((assignment: any) => (
                                <StudentMatrixRow
                                    key={assignment.id}
                                    assignment={assignment}
                                    groups={groups}
                                    subject={cls.Subject}
                                    classYear={cls.year}
                                    commonGroups={commonGroups}
                                    wrapperTemplate={wrapperTemplate}
                                />
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </main>
  );
}
