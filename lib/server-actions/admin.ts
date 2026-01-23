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

export async function updatePupil(admissionNumber: string, data: {
  firstName?: string
  lastName?: string
  gender?: string
  isActive?: boolean
}) {
  const session = await getServerSession(authOptions)
  if (!isAdmin(session?.user as any)) {
    throw new Error("Unauthorized")
  }

  try {
    await (prisma as any).pupil.update({
      where: { admissionNumber },
      data
    })

    revalidatePath("/admin")
    return { success: true }
  } catch (error) {
    console.error("Failed to update pupil:", error)
    return { success: false, error: "Failed to update pupil" }
  }
}

export async function getPupils(query: string = "") {
  const session = await getServerSession(authOptions)
  if (!isAdmin(session?.user as any)) {
    throw new Error("Unauthorized")
  }

  try {
    const pupils = await (prisma as any).pupil.findMany({
      where: {
        OR: [
          { firstName: { contains: query } },
          { lastName: { contains: query } },
          { admissionNumber: { contains: query } }
        ]
      },
      orderBy: {
        lastName: 'asc'
      }
    })
    return { success: true, pupils }
  } catch (error) {
    console.error("Failed to fetch pupils:", error)
    return { success: false, error: "Failed to fetch pupils" }
  }
}

export async function processPupilUpload(content: string) {
  const session = await getServerSession(authOptions)
  if (!isAdmin(session?.user as any)) {
    throw new Error("Unauthorized")
  }

  try {
    const lines = content.split('\n').filter(line => line.trim().startsWith('|'))
    // Skip header and separator
    const dataLines = lines.slice(2)

    const uploadedPupils = dataLines.map(line => {
      const parts = line.split('|').map(p => p.trim()).filter(Boolean)
      if (parts.length < 4) return null
      return {
        admissionNumber: parts[0],
        lastName: parts[1],
        firstName: parts[2],
        gender: parts[3]
      }
    }).filter(Boolean) as { admissionNumber: string, lastName: string, firstName: string, gender: string }[]

    if (uploadedPupils.length === 0) {
      return { success: false, error: "No valid pupil data found" }
    }

    const uploadedAdmissionNumbers = uploadedPupils.map(p => p.admissionNumber)

    // Use a transaction for consistency
    await (prisma as any).$transaction(async (tx: any) => {
      // Get existing pupils in this batch
      const existingPupils = await tx.pupil.findMany({
        where: { admissionNumber: { in: uploadedAdmissionNumbers } },
        select: { admissionNumber: true }
      })
      const existingNumbers = new Set(existingPupils.map((p: any) => p.admissionNumber))

      for (const pupil of uploadedPupils) {
        if (!existingNumbers.has(pupil.admissionNumber)) {
          // New pupil - Add
          await tx.pupil.create({
            data: {
              admissionNumber: pupil.admissionNumber,
              firstName: pupil.firstName,
              lastName: pupil.lastName,
              gender: pupil.gender,
              isActive: true
            }
          })
        } else {
          // Matching pupil exists - "do nothing" PII-wise, but ensure they are active
          await tx.pupil.update({
            where: { admissionNumber: pupil.admissionNumber },
            data: { isActive: true }
          })
        }
      }

      // Mark pupils not in upload as inactive
      await tx.pupil.updateMany({
        where: {
          admissionNumber: { notIn: uploadedAdmissionNumbers }
        },
        data: {
          isActive: false
        }
      })
    })

    revalidatePath("/admin")
    return { success: true, count: uploadedPupils.length }
  } catch (error) {
    console.error("Failed to process pupil upload:", error)
    return { success: false, error: "Failed to process upload" }
  }
}
