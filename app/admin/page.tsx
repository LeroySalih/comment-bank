import { prisma } from "@/lib/prisma"
import { UserTable } from "./_components/user-table"
import { CreateUserForm } from "./_components/create-user-form"
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

  return (
    <div className="container mx-auto py-10">
      <div className="flex justify-between items-center mb-8">
        <h1 className="text-3xl font-bold">Admin Dashboard</h1>
      </div>
      <CreateUserForm roles={roles} />
      <div className="bg-white shadow rounded-lg p-6">
        <h2 className="text-xl font-semibold mb-4">User Management</h2>
        <UserTable users={users} availableRoles={roles} />
      </div>
    </div>
  )
}
