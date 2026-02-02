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

  // Get distinct forms for pupils the user has access to
  const formsData = await prisma.pupil.findMany({
    where: {
      form: { not: null },
      isActive: true,
      Assignment: {
        some: userIsAdmin ? {} : {
          Class: userIsHoD
            ? { Subject: { User: { some: { id: userId } } } }
            : { User: { some: { id: userId } } }
        }
      }
    },
    select: {
      form: true,
    },
    distinct: ['form'],
    orderBy: { form: 'asc' }
  });

  // Get pupil counts per form
  const formCounts = await prisma.pupil.groupBy({
    by: ['form'],
    where: {
      form: { not: null },
      isActive: true,
      Assignment: {
        some: userIsAdmin ? {} : {
          Class: userIsHoD
            ? { Subject: { User: { some: { id: userId } } } }
            : { User: { some: { id: userId } } }
        }
      }
    },
    _count: { admissionNumber: true }
  });

  const forms = formsData.map(f => ({
    form: f.form!,
    pupilCount: formCounts.find(fc => fc.form === f.form)?._count.admissionNumber || 0
  }));

  // Filter logic:
  // Admin sees all subjects.
  // Others see subjects they are assigned to (via Subject or Class).
  const subjects = await prisma.subject.findMany({
    where: userIsAdmin ? {} : {
      OR: [
        { User: { some: { id: userId } } },
        { Class: { some: { User: { some: { id: userId } } } } }
      ]
    },
    include: {
      Class: {
        where: userIsAdmin ? {} : {
          // If assigned to Subject, see all classes.
          // If assigned to Class only, see only that class.
          // This nested where is tricky because we can't reference parent subject assignment easily in nested where.
          // Correct logic: return class IF (user in Subject.users) OR (user in Class.teachers).
          // Prisma doesn't support "parent" reference in nested include filter easily.
          // However, if I fetch ALL classes, I can filter in memory? Or keep it simple.
          // If I am a Teacher assigned to 1 class, I shouldn't see other classes in the subject?
          // "If a user is linked to a specific Class but not the Subject, they only have access to that specific class."
          // So I MUST filter classes.
          // BUT if I am linked to Subject, I see ALL classes.
          // Prisma query for this:
          OR: [
             { Subject: { User: { some: { id: userId } } } },
             { User: { some: { id: userId } } }
          ]
        },
        orderBy: { name: 'asc' },
        include: {
          Assignment: {
            select: {
              PupilCode: {
                select: { id: true }
              }
            }
          }
        }
      }
    },
    orderBy: { code: 'asc' }
  });

  // Calculate aggregated stats
  let totalStudents = 0;
  let totalAssignments = 0;
  let startedAssignments = 0;

  subjects.forEach((sub: any) => {
    sub.Class.forEach((cls: any) => {
      totalStudents += cls.Assignment.length;
      totalAssignments += cls.Assignment.length;
      startedAssignments += cls.Assignment.filter((a: any) => a.PupilCode.length > 0).length;
    });
  });

  const completionRate = totalAssignments > 0 ? Math.round((startedAssignments / totalAssignments) * 100) : 0;

  // Get next upcoming deadline
  const nextDeadline = await prisma.deadline.findFirst({
    where: {
      isActive: true,
      date: { gte: new Date() }
    },
    orderBy: { date: 'asc' }
  });

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark">
      {/* PageHeading */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-8">
        <div className="flex min-w-72 flex-col gap-1">
          <p className="text-slate-900 dark:text-white text-3xl font-extrabold leading-tight tracking-tight">Teacher Dashboard</p>
          <p className="text-slate-500 dark:text-slate-400 text-base font-normal leading-normal">Manage your subjects and track report completion progress.</p>
        </div>
        <div className="flex gap-3">

        </div>
      </div>

      {/* Statistics Bar */}
      <div className="px-8 pb-4 grid grid-cols-1 md:grid-cols-4 gap-4">
        <div className="bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 flex flex-col shadow-sm">
          <span className="text-slate-500 dark:text-slate-400 text-xs font-semibold uppercase tracking-wider mb-2">Total Students</span>
          <span className="text-2xl font-extrabold text-slate-900 dark:text-white">{totalStudents}</span>
        </div>
        <div className="bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 flex flex-col shadow-sm">
          <span className="text-slate-500 dark:text-slate-400 text-xs font-semibold uppercase tracking-wider mb-2">Completion Rate</span>
          <span className="text-2xl font-extrabold text-primary">{completionRate}%</span>
        </div>
        <div className="bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 flex flex-col shadow-sm">
          <span className="text-slate-500 dark:text-slate-400 text-xs font-semibold uppercase tracking-wider mb-2">Avg. Grade</span>
          <span className="text-2xl font-extrabold text-slate-900 dark:text-white">--</span>
        </div>
        <div className="bg-white dark:bg-slate-900 p-5 rounded-2xl border border-slate-200 dark:border-slate-800 flex flex-col shadow-sm">
          <span className="text-slate-500 dark:text-slate-400 text-xs font-semibold uppercase tracking-wider mb-2">Next Deadline</span>
          {nextDeadline ? (
            <div>
              <span className="text-2xl font-extrabold text-red-500">
                {new Date(nextDeadline.date).toLocaleDateString('en-GB', {
                  day: 'numeric',
                  month: 'short'
                })}
              </span>
              <p className="text-sm text-slate-600 dark:text-slate-400 mt-1 truncate" title={nextDeadline.title}>
                {nextDeadline.title}
              </p>
            </div>
          ) : (
            <span className="text-2xl font-extrabold text-slate-400">None</span>
          )}
        </div>
      </div>

      {/* SectionHeader */}
      <div className="px-8 pt-4">
        <h2 className="text-slate-900 dark:text-white text-xl font-bold leading-tight tracking-tight border-b border-slate-200 dark:border-slate-800 pb-3">My Subjects</h2>
      </div>

      {/* ImageGrid / Subjects Cards */}
      <div className="grid grid-cols-1 lg:grid-cols-2 2xl:grid-cols-3 gap-6 p-8">
{subjects.map((subject, index) => {
          // Generate a gradient based on index to vary the look
          const gradients = [
            "from-blue-600 to-indigo-700",
            "from-emerald-500 to-teal-600",
            "from-purple-500 to-indigo-600",
            "from-orange-500 to-red-600",
            "from-pink-500 to-rose-600"
          ];
          const gradient = gradients[index % gradients.length];
          const subjectTotalStudents = subject.Class.reduce((acc: any, cls: any) => acc + cls.Assignment.length, 0);

          return (
            <div key={subject.id} className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 overflow-hidden shadow-sm hover:shadow-md transition-shadow">
              <div className={`h-32 bg-gradient-to-r ${gradient} relative p-6 flex items-end`}>
                <div className="absolute top-4 right-4 bg-white/20 backdrop-blur-md rounded-full px-3 py-1 text-white text-xs font-bold">{subject.Class.length} Classes</div>
                <div className="flex items-center gap-3">
                  <div className="size-10 bg-white rounded-xl flex items-center justify-center text-slate-700 shadow-lg">
                    <span className="material-symbols-outlined">auto_stories</span>
                  </div>
                  <div className="text-white">
                    <h3 className="text-lg font-bold leading-none">{subject.title || subject.code}</h3>
                    <p className="text-blue-100 text-xs mt-1">{subjectTotalStudents} Students Total</p>
                  </div>
                </div>
              </div>
              
              <div className="p-5 space-y-4">
                <div className="space-y-3">
                  {subject.Class.length > 0 ? subject.Class.map((cls: any) => {
                     const clsTotal = cls.Assignment.length;
                     const clsStarted = cls.Assignment.filter((a: any) => a.PupilCode.length > 0).length;
                     const percent = clsTotal > 0 ? Math.round((clsStarted / clsTotal) * 100) : 0;
                     let statusColor = "text-primary";
                     let statusBg = "bg-primary";
                     let statusText = "On Track";
                     let badgeBg = "bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-400";
                     
                     if (percent === 100) {
                        statusColor = "text-emerald-500";
                        statusBg = "bg-emerald-500";
                        statusText = "Completed";
                        badgeBg = "bg-emerald-100 dark:bg-emerald-900/30 text-emerald-700 dark:text-emerald-400";
                     } else if (percent < 50) {
                        statusColor = "text-orange-500";
                        statusBg = "bg-orange-500";
                        statusText = "In Progress";
                        badgeBg = "bg-orange-100 dark:bg-orange-900/30 text-orange-700 dark:text-orange-400";
                     }
                     
                     // Only show class links for teachers
                     if (!userIsTeacher) {
                       return (
                         <div key={cls.id} className="block p-2 rounded-lg -mx-2">
                           <div className="flex flex-col gap-2">
                             <div className="flex justify-between items-center text-sm">
                               <span className="font-bold text-slate-700 dark:text-slate-300">{cls.name}</span>
                               <span className={`${statusColor} font-bold`}>{percent}%</span>
                             </div>
                             <div className="w-full bg-slate-100 dark:bg-slate-800 h-2 rounded-full overflow-hidden">
                               <div className={`${statusBg} h-full rounded-full`} style={{ width: `${percent}%` }}></div>
                             </div>
                             <div className="flex justify-between items-center text-[10px] text-slate-500 uppercase tracking-widest font-bold">
                               <span>{clsStarted}/{clsTotal} Reports Done</span>
                               <span className={`px-2 py-0.5 rounded-full ${badgeBg}`}>{statusText}</span>
                             </div>
                           </div>
                         </div>
                       );
                     }
                     
                     return (
                      <Link key={cls.id} href={`/class/${cls.id}`} className="block hover:bg-slate-50 dark:hover:bg-slate-800/50 p-2 rounded-lg transition-colors -mx-2">
                        <div className="flex flex-col gap-2">
                          <div className="flex justify-between items-center text-sm">
                            <span className="font-bold text-slate-700 dark:text-slate-300">{cls.name}</span>
                            <span className={`${statusColor} font-bold`}>{percent}%</span>
                          </div>
                          <div className="w-full bg-slate-100 dark:bg-slate-800 h-2 rounded-full overflow-hidden">
                            <div className={`${statusBg} h-full rounded-full`} style={{ width: `${percent}%` }}></div>
                          </div>
                          <div className="flex justify-between items-center text-[10px] text-slate-500 uppercase tracking-widest font-bold">
                            <span>{clsStarted}/{clsTotal} Reports Done</span>
                            <span className={`px-2 py-0.5 rounded-full ${badgeBg}`}>{statusText}</span>
                          </div>
                        </div>
                      </Link>
                     )
                  }) : (
                    <p className="text-gray-400 text-sm italic">No classes assigned.</p>
                  )}
                </div>
              </div>
            </div>
          );
        })}
      </div>

      {/* Forms Section Header */}
      {forms.length > 0 && (
        <>
          <div className="px-8 pt-4">
            <h2 className="text-slate-900 dark:text-white text-xl font-bold leading-tight tracking-tight border-b border-slate-200 dark:border-slate-800 pb-3">My Forms</h2>
          </div>

          {/* Forms Cards */}
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-4 p-8">
            {forms.map((form, index) => {
              const formGradients = [
                "from-cyan-500 to-blue-600",
                "from-violet-500 to-purple-600",
                "from-amber-500 to-orange-600",
                "from-rose-500 to-pink-600",
                "from-teal-500 to-emerald-600",
                "from-indigo-500 to-blue-600"
              ];
              const formGradient = formGradients[index % formGradients.length];

              return (
                <Link
                  key={form.form}
                  href={`/forms/${encodeURIComponent(form.form)}`}
                  className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 overflow-hidden shadow-sm hover:shadow-md transition-shadow group"
                >
                  <div className={`h-16 bg-gradient-to-r ${formGradient} flex items-center justify-center`}>
                    <span className="material-symbols-outlined text-white text-3xl group-hover:scale-110 transition-transform">groups</span>
                  </div>
                  <div className="p-4 text-center">
                    <h3 className="text-lg font-bold text-slate-900 dark:text-white">{form.form}</h3>
                    <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">{form.pupilCount} Pupils</p>
                  </div>
                </Link>
              );
            })}
          </div>
        </>
      )}
    </main>
  );
}
