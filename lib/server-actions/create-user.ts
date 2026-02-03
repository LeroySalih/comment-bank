"use server"

import { prisma } from "@/lib/prisma"
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
    const existingUser = await prisma.user.findUnique({
      where: { username }
    })

    if (existingUser) {
      return { success: false, error: "Username already exists" }
    }

    const hashedPassword = await hash(password, 12)

    const newUser = await prisma.user.create({
      data: {
        id: createId(),
        username,
        password: hashedPassword,
        Role: role ? {
          connect: { name: role }
        } : undefined
      }
    })

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
