import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';
import FormPupilRow from './_components/FormPupilRow';

export default async function FormPage({ params }: { params: Promise<{ formId: string }> }) {
  const { formId } = await params;
  const decodedFormId = decodeURIComponent(formId);
  const session = await getServerSession(authOptions);

  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userId = session?.user?.id;

  // Get classes the user is directly assigned to (for edit permissions)
  const userAssignedClasses = userId ? await prisma.class.findMany({
    where: {
      User: { some: { id: userId } }
    },
    select: { id: true }
  }) : [];
  const userAssignedClassIds = userIsAdmin
    ? [] // Admin can edit all - we'll handle this specially
    : userAssignedClasses.map(c => c.id);

  // Fetch pupils in this form - user must have access to at least one of their classes to see the pupil
  // But once they can see the pupil, they see ALL assignments (edit permission is per-class)
  const pupils = await prisma.pupil.findMany({
    where: {
      form: decodedFormId,
      isActive: true,
      Assignment: {
        some: userIsAdmin ? {} : {
          Class: userIsHoD
            ? { Subject: { User: { some: { id: userId } } } }
            : { User: { some: { id: userId } } }
        }
      }
    },
    include: {
      // Include ALL assignments - edit permission is controlled by userAssignedClassIds
      Assignment: {
        include: {
          Class: {
            include: {
              Subject: {
                include: {
                  CommentGroup: {
                    orderBy: { displayOrder: 'asc' },
                    include: {
                      CommentOption: true
                    }
                  }
                }
              }
            }
          },
          PupilCode: {
            include: {
              CommentGroup: {
                include: {
                  CommentOption: true
                }
              }
            },
            orderBy: {
              CommentGroup: { displayOrder: 'asc' }
            }
          }
        }
      }
    }
  });

  // If no pupils found, show not found or access denied
  if (pupils.length === 0) {
    return (
      <main className="flex-1 flex flex-col items-center justify-center min-h-screen bg-background-light dark:bg-background-dark">
        <div className="text-center">
          <h1 className="text-2xl font-bold text-slate-900 dark:text-white mb-2">Form Not Found</h1>
          <p className="text-slate-500 dark:text-slate-400 mb-4">No pupils found in form {decodedFormId} or you don&apos;t have access.</p>
          <Link href="/" className="text-primary hover:underline">Return to Dashboard</Link>
        </div>
      </main>
    );
  }

  // Decrypt pupil names and sort
  const decryptedPupils = pupils.map(pupil => ({
    ...pupil,
    firstName: decrypt(pupil.firstName),
    lastName: decrypt(pupil.lastName)
  }));

  decryptedPupils.sort((a, b) =>
    a.lastName.localeCompare(b.lastName) || a.firstName.localeCompare(b.firstName)
  );

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-slate-50 dark:bg-slate-950">
      {/* Page Heading */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-6 bg-white dark:bg-[#1a222c] border-b border-slate-200 dark:border-[#2d3748]">
        <div className="flex flex-col gap-1">
          <div className="flex items-center gap-2">
            <Link href="/" className="text-gray-500 hover:text-gray-900 dark:hover:text-white transition-colors">
              <ArrowLeft className="w-5 h-5" />
            </Link>
            <h1 className="text-[#111418] dark:text-white text-3xl font-black leading-tight tracking-[-0.033em]">
              Form: {decodedFormId}
            </h1>
          </div>
          <p className="text-[#617289] dark:text-gray-400 text-sm font-normal leading-normal">
            {decryptedPupils.length} Pupils
          </p>
        </div>
      </div>

      {/* Pupils Grid */}
      <div className="p-6">
        <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 overflow-hidden">
          {/* Header Row */}
          <div className="grid grid-cols-[200px_1fr_1fr] gap-4 px-4 py-3 bg-slate-100 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 text-xs font-semibold text-slate-600 dark:text-slate-400 uppercase tracking-wider">
            <div>Pupil</div>
            <div>Subjects & Codes</div>
            <div>Comments</div>
          </div>

          {/* Pupil Rows */}
          <div className="divide-y divide-slate-100 dark:divide-slate-800">
            {decryptedPupils.map((pupil) => {
              // For admin users, get all class IDs from the pupil's assignments
              const editableClassIds = userIsAdmin
                ? pupil.Assignment.map(a => a.Class.id)
                : userAssignedClassIds;

              return (
                <FormPupilRow
                  key={pupil.admissionNumber}
                  pupil={pupil}
                  userAssignedClassIds={editableClassIds}
                />
              );
            })}
          </div>
        </div>
      </div>
    </main>
  );
}
