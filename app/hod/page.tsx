
import { pool } from "@/lib/db"
import Link from "next/link"
import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { redirect } from "next/navigation"
import { decrypt } from "@/lib/encryption"

export const dynamic = 'force-dynamic'

export default async function HoDDashboard() {
  const session = await getServerSession(authOptions)
  if (!session?.user) {
    redirect("/login")
  }

  const allowAll = session.user.roles?.some((r: any) => r === 'admin' || r?.name === 'admin');

  // Fetch subjects with counts
  let subjectsQuery: string
  let subjectsParams: unknown[]

  if (allowAll) {
    subjectsQuery = `
      SELECT s.*,
             (SELECT COUNT(*) FROM "Class" c WHERE c."subjectId" = s.id) as class_count,
             (SELECT COUNT(*) FROM "CommentGroup" cg WHERE cg."subjectId" = s.id) as group_count
      FROM "Subject" s
      ORDER BY s.code ASC
    `
    subjectsParams = []
  } else {
    subjectsQuery = `
      SELECT s.*,
             (SELECT COUNT(*) FROM "Class" c WHERE c."subjectId" = s.id) as class_count,
             (SELECT COUNT(*) FROM "CommentGroup" cg WHERE cg."subjectId" = s.id) as group_count
      FROM "Subject" s
      JOIN "_SubjectToUser" su ON su."A" = s.id
      WHERE su."B" = $1
      ORDER BY s.code ASC
    `
    subjectsParams = [session.user.id]
  }

  const { rows: subjectRows } = await pool.query(subjectsQuery, subjectsParams)

  const subjects = subjectRows.map((s: any) => ({
    ...s,
    _count: {
      Class: Number(s.class_count),
      CommentGroup: Number(s.group_count)
    }
  }))

  // Get classes with assignments needing review
  const subjectIds = subjects.map((s: any) => s.id)

  let classesWithPendingReviews: any[] = []

  if (subjectIds.length > 0) {
    // Fetch classes in these subjects that have pending reviews
    const { rows: classRows } = await pool.query(
      `SELECT c.*,
              s.id as s_id, s.code as s_code, s.title as s_title,
              (SELECT COUNT(*) FROM "Assignment" a WHERE a."classId" = c.id AND a."checkStatus" = 'required_check') as pending_count
       FROM "Class" c
       JOIN "Subject" s ON s.id = c."subjectId"
       WHERE c."subjectId" = ANY($1::text[])
         AND EXISTS (
           SELECT 1 FROM "Assignment" a
           WHERE a."classId" = c.id AND a."checkStatus" = 'required_check'
         )
       ORDER BY c.name ASC`,
      [subjectIds]
    )

    // For each class, fetch up to 5 pending assignments with pupils
    for (const cls of classRows) {
      const { rows: assignmentRows } = await pool.query(
        `SELECT a.*, p."admissionNumber" as "pupil_admissionNumber",
                p."firstName" as "pupil_firstName", p."lastName" as "pupil_lastName",
                p.gender as "pupil_gender", p."isActive" as "pupil_isActive", p.form as "pupil_form"
         FROM "Assignment" a
         JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
         WHERE a."classId" = $1 AND a."checkStatus" = 'required_check'
         LIMIT 5`,
        [cls.id]
      )

      classesWithPendingReviews.push({
        id: cls.id,
        name: cls.name,
        year: cls.year,
        subjectId: cls.subjectId,
        Subject: { id: cls.s_id, code: cls.s_code, title: cls.s_title },
        _count: { Assignment: Number(cls.pending_count) },
        Assignment: assignmentRows.map((a: any) => ({
          ...a,
          Pupil: {
            admissionNumber: a.pupil_admissionNumber,
            firstName: a.pupil_firstName,
            lastName: a.pupil_lastName,
            gender: a.pupil_gender,
            isActive: a.pupil_isActive,
            form: a.pupil_form,
          }
        }))
      })
    }
  }

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark min-h-[calc(100vh-64px)]">
      {/* Page Heading */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-8 border-b border-[#f0f2f4] dark:border-gray-800 bg-white dark:bg-gray-900">
        <div className="flex min-w-72 flex-col gap-1">
          <h1 className="text-[#111418] dark:text-white text-3xl font-black leading-tight tracking-[-0.033em]">Head of Department Dashboard</h1>
          <p className="text-[#617289] dark:text-gray-400 text-base font-normal leading-normal">Manage subjects, classes, and comment banks.</p>
        </div>
        <div className="flex gap-3">
            {allowAll && (
                <button className="flex items-center justify-center rounded-xl h-10 px-4 bg-primary text-white text-sm font-bold shadow-lg shadow-primary/20 hover:bg-primary/90 transition-all">
                    <span className="material-symbols-outlined mr-2 text-lg">add</span>
                    New Subject
                </button>
            )}
        </div>
      </div>

      {/* Pending Reviews Section */}
      {classesWithPendingReviews.length > 0 && (
        <div className="p-8 border-b border-[#f0f2f4] dark:border-gray-800">
          <div className="flex items-center gap-3 mb-6">
            <div className="size-10 rounded-lg bg-amber-100 dark:bg-amber-900/30 flex items-center justify-center">
              <span className="material-symbols-outlined text-amber-600 dark:text-amber-400 text-2xl">rate_review</span>
            </div>
            <div>
              <h2 className="text-xl font-bold text-[#111418] dark:text-white">Comments Needing Review</h2>
              <p className="text-sm text-[#617289] dark:text-gray-400">
                {classesWithPendingReviews.reduce((sum: number, c: any) => sum + c._count.Assignment, 0)} comments across {classesWithPendingReviews.length} {classesWithPendingReviews.length === 1 ? 'class' : 'classes'}
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {classesWithPendingReviews.map((cls: any) => (
              <div
                key={cls.id}
                className="bg-white dark:bg-gray-900 rounded-xl border border-amber-200 dark:border-amber-800 shadow-sm overflow-hidden"
              >
                <div className="p-4 bg-amber-50 dark:bg-amber-900/20 border-b border-amber-200 dark:border-amber-800">
                  <div className="flex items-center justify-between">
                    <div>
                      <h3 className="font-bold text-[#111418] dark:text-white">{cls.name}</h3>
                      <p className="text-xs text-[#617289] dark:text-gray-400">{cls.Subject.title}</p>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="bg-amber-500 text-white text-xs font-bold px-2 py-1 rounded-full">
                        {cls._count.Assignment} pending
                      </span>
                    </div>
                  </div>
                </div>
                <div className="p-4">
                  <p className="text-xs font-medium text-[#617289] dark:text-gray-500 uppercase tracking-wider mb-2">Students awaiting review:</p>
                  <ul className="space-y-1 mb-4">
                    {cls.Assignment.slice(0, 5).map((assignment: any) => (
                      <li key={assignment.id}>
                        <Link
                          href={`/student/${assignment.id}`}
                          className="flex items-center justify-between text-sm text-[#111418] dark:text-gray-300 hover:text-primary transition-colors group/link"
                        >
                          <span>{decrypt(assignment.Pupil.firstName)} {decrypt(assignment.Pupil.lastName)}</span>
                          <span className="material-symbols-outlined text-gray-300 dark:text-gray-600 group-hover/link:text-primary text-sm transition-colors">chevron_right</span>
                        </Link>
                      </li>
                    ))}
                    {cls._count.Assignment > 5 && (
                      <li className="text-xs text-[#617289] dark:text-gray-500 italic">
                        +{cls._count.Assignment - 5} more...
                      </li>
                    )}
                  </ul>
                  <Link
                    href={`/class/${cls.id}`}
                    className="flex items-center justify-center gap-2 w-full py-2 bg-amber-100 dark:bg-amber-900/30 text-amber-700 dark:text-amber-400 rounded-lg text-sm font-medium hover:bg-amber-200 dark:hover:bg-amber-900/50 transition-colors"
                  >
                    <span className="material-symbols-outlined text-lg">visibility</span>
                    View Class
                  </Link>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Subjects Section */}
      <div className="p-8">
        <div className="flex items-center gap-3 mb-6">
          <div className="size-10 rounded-lg bg-blue-50 dark:bg-blue-900/20 flex items-center justify-center">
            <span className="material-symbols-outlined text-primary text-2xl">folder</span>
          </div>
          <div>
            <h2 className="text-xl font-bold text-[#111418] dark:text-white">Your Subjects</h2>
            <p className="text-sm text-[#617289] dark:text-gray-400">Manage classes and comment banks</p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
          {subjects.map((subject: any) => (
             <Link
               key={subject.id}
               href={`/hod/subject/${subject.id}`}
               className="group flex flex-col bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm hover:shadow-md transition-all overflow-hidden"
             >
                <div className="p-6 flex-1 flex flex-col">
                    <div className="flex justify-between items-start mb-4">
                        <div className="size-10 rounded-lg bg-blue-50 dark:bg-blue-900/20 flex items-center justify-center text-primary">
                            <span className="material-symbols-outlined text-2xl">folder</span>
                        </div>
                        <span className="material-symbols-outlined text-gray-300 dark:text-gray-600 group-hover:text-primary transition-colors">arrow_forward</span>
                    </div>

                    <h2 className="text-xl font-bold text-[#111418] dark:text-white mb-1 group-hover:text-primary transition-colors">{subject.code}</h2>
                    <h3 className="text-sm font-medium text-[#617289] dark:text-gray-400 mb-4">{subject.title}</h3>

                    <p className="text-sm text-gray-500 line-clamp-2 mb-6 flex-grow">{subject.studiedComment || "No introduction set."}</p>

                    <div className="flex gap-4 pt-4 border-t border-[#f0f2f4] dark:border-gray-800">
                        <div className="flex flex-col">
                            <span className="text-xs font-bold text-[#617289] dark:text-gray-500 uppercase tracking-wider">Classes</span>
                            <span className="text-lg font-bold text-[#111418] dark:text-white">{subject._count.Class}</span>
                        </div>
                        <div className="flex flex-col">
                            <span className="text-xs font-bold text-[#617289] dark:text-gray-500 uppercase tracking-wider">Groups</span>
                            <span className="text-lg font-bold text-[#111418] dark:text-white">{subject._count.CommentGroup}</span>
                        </div>
                    </div>
                </div>
             </Link>
          ))}

          {subjects.length === 0 && (
            <div className="col-span-full py-16 flex flex-col items-center justify-center text-center bg-white dark:bg-gray-900 rounded-xl border-2 border-dashed border-[#f0f2f4] dark:border-gray-800">
              <div className="size-16 bg-gray-50 dark:bg-gray-800 rounded-full flex items-center justify-center mb-4">
                  <span className="material-symbols-outlined text-gray-400 text-3xl">sentiment_dissatisfied</span>
              </div>
              <h3 className="text-lg font-bold text-[#111418] dark:text-white mb-1">No subjects found</h3>
              <p className="text-[#617289] dark:text-gray-400 max-w-sm mx-auto">You haven&apos;t been assigned to any subjects yet.</p>
            </div>
          )}
        </div>
      </div>
    </main>
  )
}
