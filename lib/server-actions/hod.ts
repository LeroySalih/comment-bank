"use server"

import { prisma } from "@/lib/prisma"
import { getServerSession } from "next-auth"
import { authOptions } from "@/app/api/auth/[...nextauth]/route"
import { isHoD, isAdmin } from "@/lib/access-control"
import { revalidatePath } from "next/cache"
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

export async function createSubject(formData: FormData) {
  await checkHoDAccess()
  const name = formData.get("name") as string
  const introduction = formData.get("introduction") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await prisma.course.create({
      data: {
        name,
        studiedComment: introduction
      }
    })
    revalidatePath("/hod")
    return { success: true }
  } catch (error) {
    console.error("Failed to create subject:", error)
    return { success: false, error: "Failed to create subject" }
  }
}

export async function updateSubject(subjectId: string, formData: FormData) {
  await checkHoDAccess()
  const name = formData.get("name") as string
  const introduction = formData.get("introduction") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await prisma.course.update({
      where: { id: subjectId },
      data: {
        name,
        studiedComment: introduction
      }
    })
    revalidatePath("/hod")
    return { success: true }
  } catch (error) {
    console.error("Failed to update subject:", error)
    return { success: false, error: "Failed to update subject" }
  }
}

export async function deleteSubject(subjectId: string) {
  await checkHoDAccess()

  try {
    await prisma.course.delete({
      where: { id: subjectId }
    })
    revalidatePath("/hod")
    return { success: true }
  } catch (error) {
    console.error("Failed to delete subject:", error)
    return { success: false, error: "Failed to delete subject" }
  }
}

export async function createClass(subjectId: string, formData: FormData) {
  await checkHoDAccess()
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await prisma.class.create({
      data: {
        name,
        courseId: subjectId
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
  await checkHoDAccess()
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await prisma.class.update({
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
  await checkHoDAccess()

  try {
    await prisma.class.delete({
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
  await checkHoDAccess()
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    // Get the current max order
    const maxOrderGroup = await prisma.commentGroup.findFirst({
      where: { courseId: subjectId },
      orderBy: { displayOrder: 'desc' },
      select: { displayOrder: true }
    })
    const nextOrder = (maxOrderGroup?.displayOrder ?? -1) + 1

    await prisma.commentGroup.create({
      data: {
        name,
        courseId: subjectId,
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
  await checkHoDAccess()
  const name = formData.get("name") as string

  if (!name) return { success: false, error: "Name is required" }

  try {
    await prisma.commentGroup.update({
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
  await checkHoDAccess()

  try {
    await prisma.commentGroup.delete({
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
  await checkHoDAccess()

  try {
    const transaction = items.map((item) => 
      prisma.commentGroup.update({
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
  await checkHoDAccess()
  const text = formData.get("text") as string
  const code = formData.get("code") as string

  if (!text || !code) return { success: false, error: "Text and Code are required" }

  try {
    await prisma.commentOption.create({
      data: {
        text,
        code,
        groupId
        // displayOrder not set here, using default 0
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
  await checkHoDAccess()
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
  await checkHoDAccess()

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
  await checkHoDAccess()

  try {
    const transaction = items.map((item) => 
      prisma.commentOption.update({
        where: { id: item.id },
        data: { displayOrder: item.order }
      })
    )

    await prisma.$transaction(transaction)
    // No path to revalidate easily passed here without subjectId, but client can trust optimistic update or we can fetch path. 
    // Actually, we should probably pass path or specific IDs. 
    // But typically drag and drop is optimistic. 
    // Let's rely on client router.refresh() or just pass the subjectId if we want revalidatePath.
    // For now, let's just return success.
    return { success: true }
  } catch (error) {
    console.error("Failed to reorder comments:", error)
    return { success: false, error: "Failed to reorder comments" }
  }
}

export async function createStudent(classId: string, formData: FormData) {
  await checkHoDAccess()
  const firstName = formData.get("firstName") as string
  const lastName = formData.get("lastName") as string
  const gender = formData.get("gender") as string
  const targetLevel = formData.get("targetLevel") as string
  const actualLevel = formData.get("actualLevel") as string

  if (!firstName || !lastName || !gender) return { success: false, error: "First Name, Last Name and Gender are required" }

  try {
    await prisma.student.create({
      data: {
        firstName,
        lastName,
        gender,
        targetLevel,
        actualLevel,
        classId
      }
    })
    revalidatePath(`/hod/class/${classId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to create student:", error)
    return { success: false, error: "Failed to create student" }
  }
}

export async function updateStudent(studentId: string, classId: string, formData: FormData) {
  await checkHoDAccess()
  const firstName = formData.get("firstName") as string
  const lastName = formData.get("lastName") as string
  const gender = formData.get("gender") as string
  const targetLevel = formData.get("targetLevel") as string
  const actualLevel = formData.get("actualLevel") as string

  if (!firstName || !lastName || !gender) return { success: false, error: "First Name, Last Name and Gender are required" }

  try {
    await prisma.student.update({
      where: { id: studentId },
      data: {
        firstName,
        lastName,
        gender,
        targetLevel,
        actualLevel
      }
    })
    revalidatePath(`/hod/class/${classId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to update student:", error)
    return { success: false, error: "Failed to update student" }
  }
}

export async function deleteStudent(studentId: string, classId: string) {
  await checkHoDAccess()

  try {
    await prisma.student.delete({
      where: { id: studentId }
    })
    revalidatePath(`/hod/class/${classId}`)
    return { success: true }
  } catch (error) {
    console.error("Failed to delete student:", error)
    return { success: false, error: "Failed to delete student" }
  }
}

export async function assignTeachersToClass(classId: string, teacherIds: string[]) {
  await checkHoDAccess()

  try {
    await prisma.class.update({
      where: { id: classId },
      data: {
        teachers: {
          set: teacherIds.map(id => ({ id }))
        }
      }
    })
    revalidatePath(`/hod/subject`)
    return { success: true }
  } catch (error) {
    console.error("Failed to assign teachers:", error)
    return { success: false, error: "Failed to assign teachers" }
  }
}


