import { prisma } from "@/lib/prisma"
import Link from "next/link"
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

  return (
    <div className="container mx-auto py-10">
      <div className="flex justify-between items-center mb-8">
        <h1 className="text-3xl font-bold">Admin Dashboard</h1>
        <SignOutButton />
      </div>
      
      <div className="mb-6">
        <Link
          href="/admin/ccg"
          className="inline-flex items-center gap-3 px-6 py-4 bg-white border border-gray-200 rounded-xl shadow-sm hover:shadow-md hover:border-primary/30 transition-all"
        >
          <span className="material-symbols-outlined text-2xl text-primary">comment</span>
          <div>
            <span className="text-sm font-bold text-gray-900 block">Common Comment Groups</span>
            <span className="text-xs text-gray-500">Manage global comment groups for all subjects</span>
          </div>
          <span className="material-symbols-outlined text-gray-400 ml-4">chevron_right</span>
        </Link>
      </div>

      <AdminTabs users={users} roles={roles} subjects={subjects} deadlines={deadlines} classes={classes} />
    </div>
  )
}
