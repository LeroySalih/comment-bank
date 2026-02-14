import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import CommentEditor from '@/components/CommentEditor';
import CommentStatusBadge from '@/components/CommentStatusBadge';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';
import { unstable_noStore as noStore } from 'next/cache';

// Disable caching to always get fresh data
export const dynamic = 'force-dynamic';
export const revalidate = 0;

export default async function StudentPage({ params }: { params: Promise<{ studentId: string }> }) {
  noStore(); // Disable all caching for this request

  const { studentId: assignmentId } = await params;
  const session = await getServerSession(authOptions);

  // Force Prisma to get fresh data by disconnecting/reconnecting
  await prisma.$disconnect();
  await prisma.$connect();

  const assignment = await prisma.assignment.findUnique({
    where: { id: assignmentId },
    include: {
      Pupil: true,
      PupilCode: true,
      CommonPupilCode: true,
      Class: {
        include: {
           User: { select: { id: true } },
           Subject: {
             include: {
                CommentGroup: {
                    include: {
                        CommentOption: true
                    },
                    orderBy: { displayOrder: 'asc' }
                }
             }
           }
        }
      }
    }
  });

  if (!assignment) {
    return <div>Assignment not found</div>;
  }

  // Authorization check
  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userIsTeacher = isTeacher(session?.user);

  // HoD can review comments (admins also have this capability)
  const canReviewComments = userIsHoD || userIsAdmin;

  if (userIsTeacher && !userIsAdmin && !userIsHoD) {
    const isAssigned = assignment.Class.User.some((t) => t.id === session?.user?.id);
    if (!isAssigned) {
      return <div>Student not found or access denied</div>;
    }
  }

  const subject = assignment.Class.Subject;
  const groups = subject.CommentGroup;

  // Fetch Common Comment Groups
  const commonGroups = await (prisma as any).commonCommentGroup.findMany({
    orderBy: { displayOrder: 'asc' },
    include: {
      CommonCommentOption: {
        orderBy: { displayOrder: 'asc' }
      }
    }
  });

  // Fetch format template
  const formatSetting = await (prisma as any).appSetting.findUnique({
    where: { key: 'comment_format_template' }
  });
  const formatTemplate = formatSetting?.value || '';

  // Fetch subject comment format
  const subjectFormat = (subject as any).commentFormat || null;

  // Decrypt pupil names
  const pupil = {
    ...assignment.Pupil,
    firstName: decrypt(assignment.Pupil.firstName),
    lastName: decrypt(assignment.Pupil.lastName)
  };

  // Create assignment with decrypted pupil names for CommentEditor
  const decryptedAssignment = {
    ...assignment,
    Pupil: pupil,
    finalComment: assignment.finalComment ? decrypt(assignment.finalComment) : null
  };

  return (
    <main className="min-h-screen flex flex-col bg-background-light dark:bg-background-dark">
      {/* Breadcrumbs */}
      <div className="px-10 py-4 flex items-center justify-between">
        <div className="flex flex-wrap gap-2 items-center">
            <Link href="/" className="text-[#617289] dark:text-gray-400 text-sm font-medium hover:text-primary transition-colors">Dashboard</Link>
            <span className="text-[#617289] dark:text-gray-600 material-symbols-outlined text-sm">chevron_right</span>
            <Link href={`/class/${assignment.classId}`} className="text-[#617289] dark:text-gray-400 text-sm font-medium hover:text-primary transition-colors">{assignment.Class.name}</Link>
            <span className="text-[#617289] dark:text-gray-600 material-symbols-outlined text-sm">chevron_right</span>
            <span className="text-[#111418] dark:text-white text-sm font-semibold">{pupil.firstName} {pupil.lastName}</span>
        </div>
        <div className="flex items-center gap-2 text-xs text-[#617289] dark:text-gray-400 italic">
            <span className="material-symbols-outlined text-xs">sync</span>
            Auto-save enabled
        </div>
      </div>

      {/* Page Header */}
      <div className="px-10 pb-6 flex flex-wrap justify-between items-end gap-3">
        <div className="flex min-w-72 flex-col gap-1">
            <div className="flex items-center gap-3">
                <h1 className="text-[#111418] dark:text-white text-4xl font-black leading-tight tracking-[-0.033em]">Edit Report: {pupil.firstName} {pupil.lastName}</h1>
                <CommentStatusBadge status={assignment.checkStatus || 'not_required'} size="md" />
            </div>
            <p className="text-[#617289] dark:text-gray-400 text-base font-normal">{subject.title} • {pupil.gender}</p>
        </div>
      </div>

      <div className="flex-1 px-10 pb-10 flex gap-8">
        <CommentEditor
          assignment={decryptedAssignment}
          subject={subject}
          groups={groups}
          isHoD={canReviewComments}
          commonGroups={commonGroups}
          formatTemplate={formatTemplate}
          subjectFormat={subjectFormat}
        />
      </div>
    </main>
  );
}
