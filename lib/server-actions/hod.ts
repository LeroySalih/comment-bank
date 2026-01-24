"use server"

import { prisma } from "@/lib/prisma"
import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { isHoD, isAdmin } from "@/lib/access-control"
import { revalidatePath } from "next/cache"
import { encrypt, decrypt } from "@/lib/encryption"
import { Session } from "next-auth"

async function checkHoDAccess() {
  const session = await getServerSession(authOptions)
  // Admin is also HoD effectively for management, or strictly HoD?
  // User request: "When i sign in as a HoD..."
  // I should safe check for either HoD or Admin to be safe, or just HoD if strict.
  // Access control utils usually check role inclusion.
  // Let's allow Admin to do HoD stuff too for easier testing/management, or strictly follows roles.
  // Given I am Admin, it's useful if Admin can access.
  if (!session?.user || (!isHoD(session.user as any) && !isAdmin(session.user as any))) {
    throw new Error("Unauthorized")
  }
  return session
}

async function checkSubjectAccess(subjectId: string) {
  const session = await checkHoDAccess()
  // Admin has access to everything
  if (isAdmin(session.user as any)) return session

  const hasAccess = await (prisma as any).subject.findFirst({
    where: {
      id: subjectId,
      users: { some: { id: session.user.id } }
    }
  })

  if (!hasAccess) {
    throw new Error("Unauthorized Access to Subject")
  }
  return session
}

async function checkClassAccess(classId: string) {
  const cls = await (prisma as any).class.findUnique({
    where: { id: classId },
    select: { subjectId: true }
  })
  
  if (!cls) throw new Error("Class not found")
  
  return checkSubjectAccess(cls.subjectId)
}

async function checkGroupAccess(groupId: string) {
  const group = await (prisma as any).commentGroup.findUnique({
    where: { id: groupId },
    select: { subjectId: true }
  })
  
  if (!group) throw new Error("Group not found")
  
  return checkSubjectAccess(group.subjectId)
}


export async function createClass(subjectId: string, formData: FormData) {
  await checkSubjectAccess(subjectId)
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await (prisma as any).class.create({
      data: {
        name,
        subjectId: subjectId
      }
    })
    revalidatePath(`/hod/subject/${subjectId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to create class:", error)
    return { success: false, error: "Failed to create class" }
  }
}

export async function updateClass(classId: string, subjectId: string, formData: FormData) {
  await checkSubjectAccess(subjectId)
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await (prisma as any).class.update({
      where: { id: classId },
      data: { name }
    })
    revalidatePath(`/hod/subject/${subjectId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to update class:", error)
    return { success: false, error: "Failed to update class" }
  }
}

export async function deleteClass(classId: string, subjectId: string) {
  await checkSubjectAccess(subjectId)

  try {
    await (prisma as any).class.delete({
      where: { id: classId }
    })
    revalidatePath(`/hod/subject/${subjectId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to delete class:", error)
    return { success: false, error: "Failed to delete class" }
  }
}

export async function createCommentGroup(subjectId: string, formData: FormData) {
  await checkSubjectAccess(subjectId)
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    const maxOrderGroup = await (prisma as any).commentGroup.findFirst({
      where: { subjectId: subjectId },
      orderBy: { displayOrder: 'desc' },
      select: { displayOrder: true }
    })
    const nextOrder = (maxOrderGroup?.displayOrder ?? -1) + 1

    await (prisma as any).commentGroup.create({
      data: {
        name,
        subjectId: subjectId,
        displayOrder: nextOrder
      }
    })
    revalidatePath(`/hod/subject/${subjectId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to create group:", error)
    return { success: false, error: "Failed to create group" }
  }
}

export async function updateCommentGroup(groupId: string, subjectId: string, formData: FormData) {
  await checkSubjectAccess(subjectId)
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await (prisma as any).commentGroup.update({
      where: { id: groupId },
      data: { name }
    })
    revalidatePath(`/hod/subject/${subjectId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to update group:", error)
    return { success: false, error: "Failed to update group" }
  }
}

export async function deleteCommentGroup(groupId: string, subjectId: string) {
  await checkSubjectAccess(subjectId)

  try {
    await (prisma as any).commentGroup.delete({
      where: { id: groupId }
    })
    revalidatePath(`/hod/subject/${subjectId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to delete group:", error)
    return { success: false, error: "Failed to delete group" }
  }
}
export async function reorderCommentGroups(subjectId: string, items: { id: string, order: number }[]) {
  await checkSubjectAccess(subjectId)

  try {
    const transaction = items.map((item) => 
      (prisma as any).commentGroup.update({
        where: { id: item.id },
        data: { displayOrder: item.order }
      })
    )

    await prisma.$transaction(transaction)
    revalidatePath(`/hod/subject/${subjectId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to reorder comment groups:", error)
    return { success: false, error: "Failed to reorder comment groups" }
  }
}

export async function createComment(groupId: string, subjectId: string, formData: FormData) {
  await checkSubjectAccess(subjectId)
  const text = formData.get("text") as string
  const code = formData.get("code") as string

  if (!text || !code) return { success: false, error: "Text and Code are required" }

  try {
    await prisma.commentOption.create({
      data: {
        text,
        code,
        groupId
      }
    })
    revalidatePath(`/hod/subject/${subjectId}/group/${groupId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to create comment:", error)
    return { success: false, error: "Failed to create comment" }
  }
}

export async function updateComment(commentId: string, subjectId: string, groupId: string, formData: FormData) {
  await checkSubjectAccess(subjectId)
  const text = formData.get("text") as string
  const code = formData.get("code") as string

  if (!text || !code) return { success: false, error: "Text and Code are required" }

  try {
    await prisma.commentOption.update({
      where: { id: commentId },
      data: {
        text,
        code
      }
    })
    revalidatePath(`/hod/subject/${subjectId}/group/${groupId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to update comment:", error)
    return { success: false, error: "Failed to update comment" }
  }
}

export async function deleteComment(commentId: string, subjectId: string, groupId: string) {
  await checkSubjectAccess(subjectId)

  try {
    await prisma.commentOption.delete({
      where: { id: commentId }
    })
    revalidatePath(`/hod/subject/${subjectId}/group/${groupId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to delete comment:", error)
    return { success: false, error: "Failed to delete comment" }
  }
}

export async function reorderComments(groupId: string, items: { id: string, order: number }[]) {
  await checkGroupAccess(groupId)

  try {
    const transaction = items.map((item) => 
      prisma.commentOption.update({
        where: { id: item.id },
        data: { displayOrder: item.order }
      })
    )

    await prisma.$transaction(transaction)
    return { success: true }
  } catch (error) {
    console.error("Failed to reorder comments:", error)
    return { success: false, error: "Failed to reorder comments" }
  }
}

export async function createAssignment(classId: string, formData: FormData) {
  await checkClassAccess(classId)
  const admissionNumber = formData.get("admissionNumber") as string
  const firstName = formData.get("firstName") as string
  const lastName = formData.get("lastName") as string
  const gender = formData.get("gender") as string
  const targetLevel = formData.get("targetLevel") as string
  const actualLevel = formData.get("actualLevel") as string

  if (!admissionNumber || !firstName || !lastName || !gender) {
    return { success: false, error: "Admission Number, Names and Gender are required" }
  }

  try {
    // Upsert Pupil first
    await (prisma as any).pupil.upsert({
      where: { admissionNumber },
      update: { 
        firstName: encrypt(firstName), 
        lastName: encrypt(lastName), 
        gender 
      },
      create: { 
        admissionNumber, 
        firstName: encrypt(firstName), 
        lastName: encrypt(lastName), 
        gender 
      }
    })

    // Create Assignment
    await (prisma as any).assignment.create({
      data: {
        pupilId: admissionNumber,
        classId,
        targetLevel,
        actualLevel
      }
    })
    revalidatePath(`/hod/class/${classId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to create assignment:", error)
    return { success: false, error: "Failed to create assignment" }
  }
}

export async function updateAssignment(assignmentId: string, classId: string, formData: FormData) {
  await checkClassAccess(classId)
  const firstName = formData.get("firstName") as string
  const lastName = formData.get("lastName") as string
  const gender = formData.get("gender") as string
  const targetLevel = formData.get("targetLevel") as string
  const actualLevel = formData.get("actualLevel") as string

  if (!firstName || !lastName || !gender) return { success: false, error: "Names and Gender are required" }

  try {
    const assignment = await (prisma as any).assignment.findUnique({
      where: { id: assignmentId },
      select: { pupilId: true }
    })

    if (!assignment) return { success: false, error: "Assignment not found" }

    // Update Pupil
    await (prisma as any).pupil.update({
      where: { admissionNumber: assignment.pupilId },
      data: { 
        firstName: encrypt(firstName), 
        lastName: encrypt(lastName), 
        gender 
      }
    })

    // Update Assignment
    await (prisma as any).assignment.update({
      where: { id: assignmentId },
      data: { targetLevel, actualLevel }
    })

    revalidatePath(`/hod/class/${classId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to update assignment:", error)
    return { success: false, error: "Failed to update assignment" }
  }
}

export async function deleteAssignment(assignmentId: string, classId: string) {
  await checkClassAccess(classId)

  try {
    await (prisma as any).assignment.delete({
      where: { id: assignmentId }
    })
    revalidatePath(`/hod/class/${classId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to delete assignment:", error)
    return { success: false, error: "Failed to delete assignment" }
  }
}

export { createAssignment as createStudent, updateAssignment as updateStudent, deleteAssignment as deleteStudent };



