"use server"

import { prisma } from "@/lib/prisma"
import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { isAdmin } from "@/lib/access-control"
import { hash } from "bcryptjs"
import { revalidatePath } from "next/cache"

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

    await prisma.user.create({
      data: {
        username,
        password: hashedPassword,
        roles: role ? {
          connect: { name: role }
        } : undefined
      }
    })

    revalidatePath("/admin")
    return { success: true }
  } catch (error) {
    console.error("Failed to create user:", error)
    return { success: false, error: "Failed to create user" }
  }
}
