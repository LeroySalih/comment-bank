import { prisma } from "@/lib/prisma"
import { notFound } from "next/navigation"
import Link from "next/link"
import { GroupForm } from "./_components/group-form"
import { ReorderableGroupList } from "./_components/ReorderableGroupList"
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

  const subject = await prisma.subject.findUnique({
    where: { id: subjectId },
    include: {
      CommentGroup: {
        orderBy: { displayOrder: 'asc' },
        include: {
          CommentOption: {
            orderBy: { displayOrder: 'asc' },
            select: { id: true, code: true, text: true, displayOrder: true }
          },
          _count: { select: { CommentOption: true } }
        }
      }
    }
  });

  if (!subject) notFound()

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

      <div className="p-8">
        {/* Comment Groups Section */}
        <section className="bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm flex flex-col overflow-hidden">
           <div className="p-6 border-b border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center bg-gray-50/50 dark:bg-gray-800/50">
                <div className="flex items-center gap-3">
                    <div className="size-8 rounded-lg bg-emerald-100 dark:bg-emerald-900/30 flex items-center justify-center text-emerald-600">
                        <span className="material-symbols-outlined text-lg">library_books</span>
                    </div>
                    <h2 className="text-lg font-bold text-[#111418] dark:text-white">Comment Bank</h2>
                </div>
                <GroupForm subjectId={subject.id} />
            </div>

            <div className="p-6">
                 <ReorderableGroupList
                    subjectId={subject.id}
                    initialGroups={subject.CommentGroup}
                  />
            </div>
        </section>
      </div>
    </main>
  )
}
