import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import CommentEditor from '@/components/CommentEditor';

export default async function StudentPage({ params }: { params: Promise<{ studentId: string }> }) {
  const { studentId } = await params;

  const student = await prisma.student.findUnique({
    where: { id: studentId },
    include: {
      class: {
        include: {
           course: {
             include: {
                commentGroups: {
                    include: {
                        options: true
                    }
                }
             }
           }
        }
      }
    }
  });

  if (!student) {
    return <div>Student not found</div>;
  }

  const course = student.class.course;
  const groups = course.commentGroups;

  return (
    <main className="min-h-screen p-6 bg-gray-100">
      <div className="max-w-7xl mx-auto h-full">
        <header className="mb-6 flex items-center justify-between">
            <div className="flex items-center gap-4">
                <Link href={`/class/${student.classId}`} className="p-2 hover:bg-white rounded-full transition-colors text-gray-500 hover:text-gray-900">
                    <ArrowLeft className="w-5 h-5" />
                </Link>
                <div>
                     <h1 className="text-2xl font-bold text-gray-900">{student.firstName} {student.lastName}</h1>
                     <p className="text-sm text-gray-500">{student.class.name} • {student.gender}</p>
                </div>
            </div>
        </header>
        
        <CommentEditor student={student} course={course} groups={groups} />
      </div>
    </main>
  );
}
