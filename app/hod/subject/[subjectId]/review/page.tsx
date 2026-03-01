import { pool } from "@/lib/db"
import { notFound } from "next/navigation"
import Link from "next/link"
import { getAssignmentsForReview, getReviewStats } from "@/lib/server-actions/comment-check"
import { decrypt } from "@/lib/encryption"
import ReviewList from "./_components/ReviewList"

export const dynamic = 'force-dynamic'

interface Props {
  params: Promise<{
    subjectId: string
  }>
  searchParams: Promise<{
    status?: string
  }>
}

export default async function ReviewPage({ params, searchParams }: Props) {
  const { subjectId } = await params
  const { status: statusFilter } = await searchParams

  // Fetch subject
  const { rows: subjectRows } = await pool.query(
    `SELECT * FROM "Subject" WHERE id = $1`,
    [subjectId]
  )

  if (subjectRows.length === 0) notFound()
  const subject = subjectRows[0]

  // Get review statistics
  const stats = await getReviewStats(subjectId)

  // Get assignments based on filter
  const validStatuses = ['not_required', 'required_check', 'checked_ok', 'checked_rejected'] as const
  const filterStatus = validStatuses.includes(statusFilter as any)
    ? statusFilter as typeof validStatuses[number]
    : undefined

  const assignments = await getAssignmentsForReview(subjectId, filterStatus)

  // Decrypt pupil names
  const decryptedAssignments = assignments.map(a => ({
    ...a,
    Pupil: {
      ...a.Pupil,
      firstName: decrypt(a.Pupil.firstName),
      lastName: decrypt(a.Pupil.lastName)
    }
  }))

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark min-h-[calc(100vh-64px)]">
      {/* Header */}
      <div className="bg-white dark:bg-gray-900 border-b border-[#f0f2f4] dark:border-gray-800 px-8 py-6">
        <Link href={`/hod/subject/${subjectId}`} className="inline-flex items-center text-sm font-medium text-[#617289] hover:text-primary transition-colors mb-4">
          <span className="material-symbols-outlined text-lg mr-1">arrow_back</span>
          Back to {subject.code}
        </Link>
        <div className="flex justify-between items-end">
          <div>
            <h1 className="text-3xl font-black text-[#111418] dark:text-white leading-tight tracking-tight">Review Comments</h1>
            <p className="text-[#617289] dark:text-gray-400 mt-1 text-lg">{subject.title}</p>
          </div>
          <div className="flex gap-4">
            <div className="text-right">
              <p className="text-sm text-[#617289] uppercase font-bold tracking-wider">Pending</p>
              <p className="text-xl font-black text-amber-600">{stats.pendingReview}</p>
            </div>
            <div className="h-10 w-px bg-gray-200 dark:bg-gray-700"></div>
            <div className="text-right">
              <p className="text-sm text-[#617289] uppercase font-bold tracking-wider">Approved</p>
              <p className="text-xl font-black text-emerald-600">{stats.approved}</p>
            </div>
            <div className="h-10 w-px bg-gray-200 dark:bg-gray-700"></div>
            <div className="text-right">
              <p className="text-sm text-[#617289] uppercase font-bold tracking-wider">Rejected</p>
              <p className="text-xl font-black text-red-600">{stats.rejected}</p>
            </div>
          </div>
        </div>
      </div>

      {/* Filter Tabs */}
      <div className="bg-white dark:bg-gray-900 border-b border-[#f0f2f4] dark:border-gray-800 px-8">
        <div className="flex gap-1">
          <Link
            href={`/hod/subject/${subjectId}/review`}
            className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
              !statusFilter
                ? 'border-primary text-primary'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Pending Review ({stats.pendingReview})
          </Link>
          <Link
            href={`/hod/subject/${subjectId}/review?status=checked_ok`}
            className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
              statusFilter === 'checked_ok'
                ? 'border-primary text-primary'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Approved ({stats.approved})
          </Link>
          <Link
            href={`/hod/subject/${subjectId}/review?status=checked_rejected`}
            className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
              statusFilter === 'checked_rejected'
                ? 'border-primary text-primary'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Rejected ({stats.rejected})
          </Link>
        </div>
      </div>

      {/* Content */}
      <div className="p-8">
        {decryptedAssignments.length === 0 ? (
          <div className="bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 p-12 text-center">
            <span className="material-symbols-outlined text-4xl text-gray-300 dark:text-gray-600 mb-4">inbox</span>
            <p className="text-gray-500 dark:text-gray-400">
              {statusFilter === 'checked_ok'
                ? 'No approved comments yet'
                : statusFilter === 'checked_rejected'
                  ? 'No rejected comments'
                  : 'No comments pending review'
              }
            </p>
          </div>
        ) : (
          <ReviewList
            assignments={decryptedAssignments}
            subjectId={subjectId}
          />
        )}
      </div>
    </main>
  )
}
