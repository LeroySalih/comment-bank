import { prisma } from "@/lib/prisma"
import { AdminTabs } from "./_components/admin-tabs"
import SignOutButton from "@/components/SignOutButton"

export const dynamic = 'force-dynamic'

export default async function AdminPage() {
  const users = await prisma.user.findMany({
    include: {
      roles: true
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
        select: { classes: true, commentGroups: true }
      },
      users: {
        select: { id: true, username: true, roles: { select: { name: true } } }
      }
    }
  })

  return (
    <div className="container mx-auto py-10">
      <div className="flex justify-between items-center mb-8">
        <h1 className="text-3xl font-bold">Admin Dashboard</h1>
        <SignOutButton />
      </div>
      
      <AdminTabs users={users} roles={roles} subjects={subjects} />
    </div>
  )
}
