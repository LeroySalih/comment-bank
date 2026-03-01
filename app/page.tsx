import { pool } from '@/lib/db'
import Link from 'next/link';
import { getServerSession } from 'next-auth';
import { authOptions } from './api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';

export default async function Home() {
  const session = await getServerSession(authOptions);

  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userIsTeacher = isTeacher(session?.user);
  const userId = session?.user?.id;

  // Get distinct forms for pupils the user has access to, with pupil counts
  let formsQuery: string;
  let formsParams: unknown[];

  if (userIsAdmin) {
    formsQuery = `
      SELECT p.form, COUNT(*) as pupil_count
      FROM "Pupil" p
      WHERE p.form IS NOT NULL AND p."isActive" = true
        AND EXISTS (
          SELECT 1 FROM "Assignment" a WHERE a."pupilId" = p."admissionNumber"
        )
      GROUP BY p.form
      ORDER BY p.form ASC
    `;
    formsParams = [];
  } else if (userIsHoD) {
    formsQuery = `
      SELECT p.form, COUNT(*) as pupil_count
      FROM "Pupil" p
      WHERE p.form IS NOT NULL AND p."isActive" = true
        AND EXISTS (
          SELECT 1 FROM "Assignment" a
          JOIN "Class" c ON c.id = a."classId"
          JOIN "Subject" s ON s.id = c."subjectId"
          JOIN "_SubjectToUser" su ON su."A" = s.id
          WHERE a."pupilId" = p."admissionNumber" AND su."B" = $1
        )
      GROUP BY p.form
      ORDER BY p.form ASC
    `;
    formsParams = [userId];
  } else {
    formsQuery = `
      SELECT p.form, COUNT(*) as pupil_count
      FROM "Pupil" p
      WHERE p.form IS NOT NULL AND p."isActive" = true
        AND EXISTS (
          SELECT 1 FROM "Assignment" a
          JOIN "Class" c ON c.id = a."classId"
          JOIN "_ClassToUser" cu ON cu."A" = c.id
          WHERE a."pupilId" = p."admissionNumber" AND cu."B" = $1
        )
      GROUP BY p.form
      ORDER BY p.form ASC
    `;
    formsParams = [userId];
  }

  const { rows: formsRows } = await pool.query(formsQuery, formsParams);
  const forms = formsRows.map((r: any) => ({
    form: r.form as string,
    pupilCount: Number(r.pupil_count)
  }));

  // Fetch subjects and their classes with assignment counts
  let subjectsQuery: string;
  let subjectsParams: unknown[];

  if (userIsAdmin) {
    subjectsQuery = `
      SELECT s.id as s_id, s.code as s_code, s.title as s_title,
             c.id as c_id, c.name as c_name, c.year as c_year,
             COUNT(DISTINCT a.id) as assignment_count,
             COUNT(DISTINCT CASE WHEN pc_count.pc_cnt > 0 THEN a.id END) as started_count
      FROM "Subject" s
      LEFT JOIN "Class" c ON c."subjectId" = s.id
      LEFT JOIN "Assignment" a ON a."classId" = c.id
      LEFT JOIN (
        SELECT "assignmentId", COUNT(*) as pc_cnt
        FROM "PupilCode"
        GROUP BY "assignmentId"
      ) pc_count ON pc_count."assignmentId" = a.id
      GROUP BY s.id, s.code, s.title, c.id, c.name, c.year
      ORDER BY s.code ASC, c.name ASC
    `;
    subjectsParams = [];
  } else {
    subjectsQuery = `
      SELECT s.id as s_id, s.code as s_code, s.title as s_title,
             c.id as c_id, c.name as c_name, c.year as c_year,
             COUNT(DISTINCT a.id) as assignment_count,
             COUNT(DISTINCT CASE WHEN pc_count.pc_cnt > 0 THEN a.id END) as started_count
      FROM "Subject" s
      LEFT JOIN "_SubjectToUser" su ON su."A" = s.id
      LEFT JOIN "Class" c ON c."subjectId" = s.id
      LEFT JOIN "_ClassToUser" cu ON cu."A" = c.id
      LEFT JOIN "Assignment" a ON a."classId" = c.id
      LEFT JOIN (
        SELECT "assignmentId", COUNT(*) as pc_cnt
        FROM "PupilCode"
        GROUP BY "assignmentId"
      ) pc_count ON pc_count."assignmentId" = a.id
      WHERE su."B" = $1 OR cu."B" = $1
      GROUP BY s.id, s.code, s.title, c.id, c.name, c.year
      ORDER BY s.code ASC, c.name ASC
    `;
    subjectsParams = [userId];
  }

  const { rows: subjectRows } = await pool.query(subjectsQuery, subjectsParams);

  // Group rows into subjects with nested classes
  const subjectMap = new Map<string, any>();
  for (const row of subjectRows) {
    if (!subjectMap.has(row.s_id)) {
      subjectMap.set(row.s_id, {
        id: row.s_id,
        code: row.s_code,
        title: row.s_title,
        Class: []
      });
    }
    const subject = subjectMap.get(row.s_id)!;
    if (row.c_id) {
      subject.Class.push({
        id: row.c_id,
        name: row.c_name,
        year: row.c_year,
        Assignment: Array(Number(row.assignment_count)).fill(null).map((_, i) => ({
          PupilCode: i < Number(row.started_count) ? [{ id: 'x' }] : []
        }))
      });
    }
  }
  const subjects = Array.from(subjectMap.values());

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
  const { rows: deadlineRows } = await pool.query(
    `SELECT * FROM "Deadline"
     WHERE "isActive" = true AND date >= NOW()
     ORDER BY date ASC
     LIMIT 1`
  );
  const nextDeadline = deadlineRows[0] ?? null;

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark">
      {/* PageHeading */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-8">
        <div className="flex min-w-72 flex-col gap-1">
          <p className="text-slate-900 dark:text-white text-3xl font-extrabold leading-tight tracking-tight">Pupil Comments</p>
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
