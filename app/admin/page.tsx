import { prisma } from "@/lib/prisma"
import { AdminTabs } from "./_components/admin-tabs"
import SignOutButton from "@/components/SignOutButton"

export const dynamic = 'force-dynamic'

export default async function AdminPage() {
  const users = await prisma.user.findMany({
    select: {
      id: true,
      username: true,
      isActive: true,
      Role: true
    },
    orderBy: {
      username: 'asc'
    }
  })

  const roles = await prisma.role.findMany({
    orderBy: {
      name: 'asc'
    }
  })

  const subjects = await (prisma as any).subject.findMany({
    orderBy: { code: 'asc' },
    include: {
      _count: {
        select: { Class: true, CommentGroup: true }
      },
      User: {
        select: { id: true, username: true, Role: { select: { name: true } } }
      },
      CommentGroup: {
        orderBy: { displayOrder: 'asc' },
        include: {
          CommentOption: {
            orderBy: { displayOrder: 'asc' }
          }
        }
      }
    }
  })

  const deadlines = await prisma.deadline.findMany({
    orderBy: { date: 'asc' }
  })

  const classes = await prisma.class.findMany({
    select: { id: true, name: true },
    orderBy: { name: 'asc' }
  })

  const commonGroups = await (prisma as any).commonCommentGroup.findMany({
    orderBy: { displayOrder: 'asc' },
    include: {
      CommonCommentOption: {
        orderBy: { displayOrder: 'asc' }
      }
    }
  })

  const formatSetting = await (prisma as any).appSetting.findUnique({
    where: { key: 'comment_format_template' }
  })
  const commentFormatTemplate = formatSetting?.value || ''

  return (
    <div className="container mx-auto py-10">
      <div className="flex justify-between items-center mb-8">
        <h1 className="text-3xl font-bold">Admin Dashboard</h1>
        <SignOutButton />
      </div>
      
      <AdminTabs users={users} roles={roles} subjects={subjects} deadlines={deadlines} classes={classes} commonGroups={commonGroups} commentFormatTemplate={commentFormatTemplate} />
    </div>
  )
}
