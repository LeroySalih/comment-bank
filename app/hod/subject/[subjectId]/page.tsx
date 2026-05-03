import { pool } from "@/lib/db"
import { notFound } from "next/navigation"
import Link from "next/link"
import { GroupForm } from "./_components/group-form"
import { ReorderableGroupList } from "./_components/ReorderableGroupList"
import { SubjectCommentFormat } from "./_components/SubjectCommentFormat"
import { AuditButton } from "./_components/audit-button"
import { countWords } from "@/lib/utils"
import { getReviewStats } from "@/lib/server-actions/comment-check"

export const dynamic = 'force-dynamic'

interface Props {
  params: Promise<{
    subjectId: string
  }>
}

export default async function SubjectPage({ params }: Props) {
  const { subjectId } = await params

  // Fetch subject
  const { rows: subjectRows } = await pool.query(
    `SELECT * FROM "Subject" WHERE id = $1`,
    [subjectId]
  )

  if (subjectRows.length === 0) notFound()
  const subjectRow = subjectRows[0]

  // Fetch comment groups with options and counts
  const { rows: groupRows } = await pool.query(
    `SELECT cg.*,
            (SELECT COUNT(*) FROM "CommentOption" co WHERE co."groupId" = cg.id) as option_count
     FROM "CommentGroup" cg
     WHERE cg."subjectId" = $1
     ORDER BY cg."displayOrder" ASC`,
    [subjectId]
  )

  if (groupRows.length > 0) {
    const groupIds = groupRows.map((g: any) => g.id)
    const { rows: optRows } = await pool.query(
      `SELECT id, code, text, "displayOrder", "groupId"
       FROM "CommentOption"
       WHERE "groupId" = ANY($1::text[])
       ORDER BY "displayOrder" ASC`,
      [groupIds]
    )
    const optsByGroup = new Map<string, any[]>()
    for (const opt of optRows) {
      const arr = optsByGroup.get(opt.groupId) ?? []
      arr.push(opt)
      optsByGroup.set(opt.groupId, arr)
    }
    for (const g of groupRows) {
      (g as any).CommentOption = optsByGroup.get(g.id) ?? []
      ;(g as any)._count = { CommentOption: Number(g.option_count) }
    }
  }

  const subject = {
    ...subjectRow,
    CommentGroup: groupRows,
  }

  // Fetch CCG group names so HODs can see which names enable overrides
  const { rows: ccgRows } = await pool.query(
    `SELECT id, name, title FROM "CommonCommentGroup" ORDER BY "displayOrder" ASC`
  )
  const ccgGroups = ccgRows as { id: string; name: string; title: string }[]

  // Get review statistics
  const reviewStats = await getReviewStats(subjectId)

  const totalWordsInSubject = subject.CommentGroup.reduce((acc: number, group: any) => {
    return acc + group.CommentOption.reduce((sum: number, opt: any) => sum + countWords(opt.text), 0)
  }, 0)
  const avgWordsPerGroup = subject.CommentGroup.length > 0 ? (totalWordsInSubject / subject.CommentGroup.length).toFixed(1) : 0

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark min-h-[calc(100vh-64px)]">
      {/* Header */}
      <div className="bg-white dark:bg-gray-900 border-b border-[#f0f2f4] dark:border-gray-800 px-8 py-6">
          <Link href="/hod" className="inline-flex items-center text-sm font-medium text-[#617289] hover:text-primary transition-colors mb-4">
              <span className="material-symbols-outlined text-lg mr-1">arrow_back</span>
              Back to Department
          </Link>
          <div className="flex justify-between items-end">
             <div>
                <h1 className="text-3xl font-black text-[#111418] dark:text-white leading-tight tracking-tight">{subject.code}</h1>
                <p className="text-[#617289] dark:text-gray-400 mt-1 text-lg">{subject.title}</p>
             </div>
             <div className="flex items-center gap-4">
                 {reviewStats.pendingReview > 0 && (
                   <Link
                     href={`/hod/subject/${subjectId}/review`}
                     className="flex items-center gap-2 px-4 py-2 bg-amber-100 dark:bg-amber-900/30 text-amber-700 dark:text-amber-400 rounded-lg hover:bg-amber-200 dark:hover:bg-amber-900/50 transition-colors font-medium text-sm"
                   >
                     <span className="material-symbols-outlined text-lg">rate_review</span>
                     Review Comments ({reviewStats.pendingReview})
                   </Link>
                 )}
                 {reviewStats.pendingReview === 0 && (
                   <Link
                     href={`/hod/subject/${subjectId}/review`}
                     className="flex items-center gap-2 px-4 py-2 bg-gray-100 dark:bg-gray-800 text-gray-600 dark:text-gray-400 rounded-lg hover:bg-gray-200 dark:hover:bg-gray-700 transition-colors font-medium text-sm"
                   >
                     <span className="material-symbols-outlined text-lg">rate_review</span>
                     Review Comments
                   </Link>
                 )}
                 <AuditButton
                   subjectId={subjectId}
                   subjectTitle={subject.title ?? subject.code}
                 />
                 <div className="h-10 w-px bg-gray-200 dark:bg-gray-700 hidden md:block"></div>
                 <div className="text-right hidden md:block">
                     <p className="text-sm text-[#617289] uppercase font-bold tracking-wider">Avg. Words</p>
                     <p className="text-xl font-black text-[#111418] dark:text-white">{avgWordsPerGroup}</p>
                 </div>
                 <div className="h-10 w-px bg-gray-200 dark:bg-gray-700 hidden md:block"></div>
                 <div className="text-right hidden md:block">
                     <p className="text-sm text-[#617289] uppercase font-bold tracking-wider">Groups</p>
                     <p className="text-xl font-black text-[#111418] dark:text-white">{subject.CommentGroup.length}</p>
                 </div>
             </div>
          </div>
      </div>

      <div className="p-8 space-y-6">
        {/* Comment Groups Section */}
        <section className="bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm flex flex-col overflow-hidden">
           <div className="p-6 border-b border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center bg-gray-50/50 dark:bg-gray-800/50">
                <div className="flex items-center gap-3">
                    <div className="size-8 rounded-lg bg-emerald-100 dark:bg-emerald-900/30 flex items-center justify-center text-emerald-600">
                        <span className="material-symbols-outlined text-lg">library_books</span>
                    </div>
                    <h2 className="text-lg font-bold text-[#111418] dark:text-white">Comment Bank</h2>
                </div>
                <GroupForm subjectId={subject.id} ccgGroups={ccgGroups} />
            </div>

            <div className="p-6">
                 <ReorderableGroupList
                    subjectId={subject.id}
                    initialGroups={subject.CommentGroup}
                    ccgGroups={ccgGroups}
                    subjectTitle={subject.title ?? subject.code}
                  />
            </div>
        </section>

        {/* Subject Comment Format */}
        <SubjectCommentFormat
          subjectId={subject.id}
          initialFormat={subject.commentFormat}
          groups={subject.CommentGroup}
        />
      </div>
    </main>
  )
}
