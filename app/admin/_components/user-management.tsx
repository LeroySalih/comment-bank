import { CreateUserForm } from "./create-user-form"
import { UserTable } from "./user-table"

interface UserManagementProps {
  users: any[]
  roles: any[]
}

export function UserManagement({ users, roles }: UserManagementProps) {
  return (
    <div className="space-y-8">
      <CreateUserForm roles={roles} />
      <div className="bg-white shadow rounded-lg p-6">
        <h2 className="text-xl font-semibold mb-4">User Management</h2>
        <UserTable users={users} availableRoles={roles} />
      </div>
    </div>
  )
}
