import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import CommentEditor from '@/components/CommentEditor';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';

export default async function StudentPage({ params }: { params: Promise<{ studentId: string }> }) {
  const { studentId: assignmentId } = await params;
  const session = await getServerSession(authOptions);

  const assignment = await ((prisma as any).assignment.findUnique({
    where: { id: assignmentId },
    include: {
      pupil: true,
      codes: true,
      class: {
        include: {
           teachers: { select: { id: true } },
           subject: {
             include: {
                commentGroups: {
                    include: {
                        options: true
                    },
                    orderBy: { displayOrder: 'asc' }
                }
             }
           }
        }
      }
    }
  }) as any);

  if (!assignment) {
    return <div>Assignment not found</div>;
  }

  // Authorization check
  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userIsTeacher = isTeacher(session?.user);

  if (userIsTeacher && !userIsAdmin && !userIsHoD) {
    const isAssigned = assignment.class.teachers.some((t: any) => t.id === session?.user?.id);
    if (!isAssigned) {
      return <div>Student not found or access denied</div>;
    }
  }

  const subject = assignment.class.subject;
  const groups = subject.commentGroups;

  // Decrypt pupil names
  assignment.pupil = {
    ...assignment.pupil,
    firstName: decrypt(assignment.pupil.firstName),
    lastName: decrypt(assignment.pupil.lastName)
  };

  return (
    <main className="min-h-screen flex flex-col bg-background-light dark:bg-background-dark">
      {/* Breadcrumbs */}
      <div className="px-10 py-4 flex items-center justify-between">
        <div className="flex flex-wrap gap-2 items-center">
            <Link href="/" className="text-[#617289] dark:text-gray-400 text-sm font-medium hover:text-primary transition-colors">Dashboard</Link>
            <span className="text-[#617289] dark:text-gray-600 material-symbols-outlined text-sm">chevron_right</span>
            <Link href={`/class/${assignment.classId}`} className="text-[#617289] dark:text-gray-400 text-sm font-medium hover:text-primary transition-colors">{assignment.class.name}</Link>
            <span className="text-[#617289] dark:text-gray-600 material-symbols-outlined text-sm">chevron_right</span>
            <span className="text-[#111418] dark:text-white text-sm font-semibold">{assignment.pupil.firstName} {assignment.pupil.lastName}</span>
        </div>
        <div className="flex items-center gap-2 text-xs text-[#617289] dark:text-gray-400 italic">
            <span className="material-symbols-outlined text-xs">sync</span>
            Auto-save enabled
        </div>
      </div>

      {/* Page Header */}
      <div className="px-10 pb-6 flex flex-wrap justify-between items-end gap-3">
        <div className="flex min-w-72 flex-col gap-1">
            <h1 className="text-[#111418] dark:text-white text-4xl font-black leading-tight tracking-[-0.033em]">Edit Report: {assignment.pupil.firstName} {assignment.pupil.lastName}</h1>
            <p className="text-[#617289] dark:text-gray-400 text-base font-normal">{subject.title} • {assignment.pupil.gender}</p>
        </div>
        <div className="flex gap-3">
            <button className="flex min-w-[100px] cursor-pointer items-center justify-center rounded-lg h-10 px-4 bg-[#f0f2f4] dark:bg-gray-800 text-[#111418] dark:text-white text-sm font-bold hover:bg-gray-200 transition-colors">
                <span className="truncate">Save Draft</span>
            </button>
            <button className="flex min-w-[140px] cursor-pointer items-center justify-center rounded-lg h-10 px-4 bg-primary text-white text-sm font-bold hover:opacity-90 transition-opacity gap-2">
                <span className="material-symbols-outlined text-base">check_circle</span>
                <span className="truncate">Mark as Complete</span>
            </button>
        </div>
      </div>

      <div className="flex-1 px-10 pb-10 flex gap-8">
        <CommentEditor assignment={assignment} subject={subject} groups={groups} />
      </div>
    </main>
  );
}
