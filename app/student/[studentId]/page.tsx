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
    <main className="min-h-screen p-6 bg-gray-100">
      <div className="max-w-7xl mx-auto h-full">
        <header className="mb-6 flex items-center justify-between">
            <div className="flex items-center gap-4">
                <Link href={`/class/${assignment.classId}`} className="p-2 hover:bg-white rounded-full transition-colors text-gray-500 hover:text-gray-900">
                    <ArrowLeft className="w-5 h-5" />
                </Link>
                <div>
                     <h1 className="text-2xl font-bold text-gray-900">{assignment.pupil.firstName} {assignment.pupil.lastName}</h1>
                     <p className="text-sm text-gray-500">{assignment.class.name} • {assignment.pupil.gender}</p>
                </div>
            </div>
        </header>
        
        <CommentEditor assignment={assignment} subject={subject} groups={groups} />
      </div>
    </main>
  );
}
