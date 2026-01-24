
import { prisma } from "@/lib/prisma"
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
  
  const group = await (prisma as any).commentGroup.findUnique({
    where: { id: groupId },
    include: {
      options: {
        orderBy: [
          { displayOrder: 'asc' },
          { code: 'asc' }
        ]
      },
      subject: true
    }
  }) as any;

  if (!group || group.subjectId !== subjectId) notFound()
  
  const totalWords = group.options.reduce((acc: any, opt: any) => acc + countWords(opt.text), 0)
  const avgWords = group.options.length > 0 ? (totalWords / group.options.length).toFixed(1) : 0

  return (
    <div className="container mx-auto py-10">
      <div className="flex flex-col gap-2 mb-8">
        <div className="flex items-center gap-2 text-sm text-gray-500">
           <Link href="/hod" className="hover:text-indigo-600">Dashboard</Link>
           <span>/</span>
           <Link href={`/hod/subject/${subjectId}`} className="hover:text-indigo-600">{group.subject.code}</Link>
           <span>/</span>
           <span>{group.name}</span>
        </div>
        <h1 className="text-3xl font-bold">{group.name} - Comments</h1>
      </div>

      <div className="bg-white shadow rounded-lg p-6">
        <div className="flex justify-between items-center mb-6">
           <div>
            <h2 className="text-xl font-semibold">Comment Options</h2>
            {group.options.length > 0 && (
              <p className="text-sm text-gray-500">{avgWords} avg words per comment</p>
            )}
           </div>
           <CommentForm groupId={group.id} subjectId={subjectId} />
        </div>

        <div className="grid grid-cols-1 gap-4">
          <CommentList 
            initialComments={group.options.map((o: any) => ({ ...o, order: o.displayOrder }))}
            subjectId={subjectId}
            groupId={groupId}
          />
          {group.options.length === 0 && (
             <p className="text-gray-500 text-center py-10">No comments added to this group yet.</p>
          )}
        </div>
      </div>
    </div>
  )
}

