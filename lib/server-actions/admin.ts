"use server"

import { prisma } from "@/lib/prisma"
import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { isAdmin } from "@/lib/access-control"
import { revalidatePath } from "next/cache"

export async function updateUserRoles(userId: string, roleNames: string[]) {
  const session = await getServerSession(authOptions)
  if (!isAdmin(session?.user as any)) { // Casting because session.user might not infer roles correctly here without global type aug pickup
    throw new Error("Unauthorized")
  }

  try {
    // Get role IDs for the names
    const roles = await prisma.role.findMany({
      where: {
        name: { in: roleNames }
      }
    })

    await prisma.user.update({
      where: { id: userId },
      data: {
        roles: {
          set: [], // Clear existing
          connect: roles.map(r => ({ id: r.id })) // Connect new
        }
      }
    })

    revalidatePath("/admin")
    return { success: true }
  } catch (error) {
    console.error("Failed to update roles:", error)
    return { success: false, error: "Failed to update roles" }
  }
}
