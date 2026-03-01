"use server"

import { pool } from "@/lib/db"
import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { isAdmin } from "@/lib/access-control"
import { hash } from "bcryptjs"
import { revalidatePath } from "next/cache"
import { createId } from "@paralleldrive/cuid2"
import { createAuditLog } from "@/lib/audit-log"

export async function createUser(formData: FormData) {
  const session = await getServerSession(authOptions)
  if (!isAdmin(session?.user as any)) {
    return { success: false, error: "Unauthorized" }
  }

  const username = formData.get("username") as string
  const password = formData.get("password") as string
  const role = formData.get("role") as string // "teacher", "hod", "admin"

  if (!username || !password) {
    return { success: false, error: "Username and password are required" }
  }

  try {
    const { rows: existingRows } = await pool.query(
      `SELECT id FROM "User" WHERE username = $1`,
      [username]
    )
    const existingUser = existingRows[0] ?? null

    if (existingUser) {
      return { success: false, error: "Username already exists" }
    }

    const hashedPassword = await hash(password, 12)

    const userId = createId()
    await pool.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)`,
      [userId, username, hashedPassword]
    )

    if (role) {
      const { rows: roleRows } = await pool.query(
        `SELECT id FROM "Role" WHERE name = $1`,
        [role]
      )
      if (roleRows.length > 0) {
        await pool.query(
          `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [roleRows[0].id, userId]
        )
      }
    }

    const newUser = { id: userId }

    // Audit log
    await createAuditLog({
      action: 'create_user',
      entityType: 'user',
      entityId: newUser.id,
      details: {
        after: { username, role: role || null }
      }
    })

    revalidatePath("/admin")
    return { success: true }
  } catch (error) {
    console.error("Failed to create user:", error)
    return { success: false, error: "Failed to create user" }
  }
}
