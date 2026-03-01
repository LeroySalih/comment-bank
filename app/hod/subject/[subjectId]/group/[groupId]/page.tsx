
import { pool } from "@/lib/db"
import { notFound } from "next/navigation"
import Link from "next/link"
import { CommentForm } from "./_components/comment-form"
import { CommentList } from "./_components/comment-list"
import { countWords } from "@/lib/utils"

export const dynamic = 'force-dynamic'

interface Props {
  params: Promise<{
    subjectId: string
    groupId: string
  }>
}

export default async function GroupPage({ params }: Props) {
  const { subjectId, groupId } = await params

  // Fetch comment group with subject
  const { rows: groupRows } = await pool.query(
    `SELECT cg.*, s.id as s_id, s.code as s_code, s.title as s_title
     FROM "CommentGroup" cg
     JOIN "Subject" s ON s.id = cg."subjectId"
     WHERE cg.id = $1`,
    [groupId]
  )

  if (groupRows.length === 0 || groupRows[0].subjectId !== subjectId) notFound()

  const groupRow = groupRows[0]

  // Fetch comment options
  const { rows: optionRows } = await pool.query(
    `SELECT * FROM "CommentOption" WHERE "groupId" = $1 ORDER BY "displayOrder" ASC, code ASC`,
    [groupId]
  )

  const group = {
    id: groupRow.id,
    name: groupRow.name,
    displayOrder: groupRow.displayOrder,
    subjectId: groupRow.subjectId,
    title: groupRow.title,
    isLinked: groupRow.isLinked,
    linkedField: groupRow.linkedField,
    CommentOption: optionRows,
    Subject: { id: groupRow.s_id, code: groupRow.s_code, title: groupRow.s_title },
  }

  const totalWords = group.CommentOption.reduce((acc: number, opt: any) => acc + countWords(opt.text), 0)
  const avgWords = group.CommentOption.length > 0 ? (totalWords / group.CommentOption.length).toFixed(1) : 0

  return (
    <div className="container mx-auto py-10">
      <div className="flex flex-col gap-2 mb-8">
        <div className="flex items-center gap-2 text-sm text-gray-500">
           <Link href="/hod" className="hover:text-indigo-600">Dashboard</Link>
           <span>/</span>
           <Link href={`/hod/subject/${subjectId}`} className="hover:text-indigo-600">{group.Subject.code}</Link>
           <span>/</span>
           <span>{group.name}</span>
        </div>
        <h1 className="text-3xl font-bold">{group.name} - Comments</h1>
      </div>

      <div className="bg-white shadow rounded-lg p-6">
        <div className="flex justify-between items-center mb-6">
           <div>
            <h2 className="text-xl font-semibold">Comment Options</h2>
            {group.CommentOption.length > 0 && (
              <p className="text-sm text-gray-500">{avgWords} avg words per comment</p>
            )}
           </div>
           <CommentForm groupId={group.id} subjectId={subjectId} />
        </div>

        <div className="grid grid-cols-1 gap-4">
          <CommentList
            initialComments={group.CommentOption.map((o: any) => ({ ...o, order: o.displayOrder }))}
            subjectId={subjectId}
            groupId={groupId}
          />
          {group.CommentOption.length === 0 && (
             <p className="text-gray-500 text-center py-10">No comments added to this group yet.</p>
          )}
        </div>
      </div>
    </div>
  )
}
