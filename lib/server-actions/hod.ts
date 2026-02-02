"use server"

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { withRole } from '@/lib/auth/with-role'
import { createId } from '@paralleldrive/cuid2'
import { subjectRepository } from '@/lib/db/repositories/subject-repository'
import { classRepository } from '@/lib/db/repositories/class-repository'
import { handleServerActionError, ForbiddenError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import {
  CreateCommentGroupSchema,
  UpdateCommentGroupSchema,
  DeleteCommentGroupSchema,
  CreateCommentOptionSchema,
  UpdateCommentOptionSchema,
  DeleteCommentOptionSchema,
  ReorderCommentGroupsSchema,
  ReorderCommentOptionsSchema,
  validateFormData
} from '@/lib/validation-schemas'
import { getServerSession } from 'next-auth'
import { authOptions } from '@/app/api/auth/[...nextauth]/route'

/**
 * Helper to check if user is HOD of a specific subject
 */
async function checkSubjectAccess(subjectId: string): Promise<void> {
  const session = await getServerSession(authOptions)
  if (!session?.user) {
    throw new ForbiddenError('Not authenticated')
  }

  // Admins have access to everything
  if (session.user.roles?.includes('admin')) {
    return
  }

  // Check if user is HOD of this subject
  const subject = await prisma.subject.findFirst({
    where: {
      id: subjectId,
      User: {
        some: { id: session.user.id }
      }
    }
  })

  if (!subject) {
    throw new ForbiddenError('You do not have access to this subject')
  }
}

// ============================================================================
// Class Management Actions
// ============================================================================

/**
 * Create a new class (HOD/Admin only)
 */
export const createClass = withRole(['admin', 'hod'], async (
  subjectId: string,
  formData: FormData
) => {
  try {
    await checkSubjectAccess(subjectId)

    const name = formData.get('name') as string
    const year = formData.get('year') as string

    if (!name) {
      return {
        success: false,
        error: 'Class name is required',
        code: 'VALIDATION_ERROR'
      }
    }

    await prisma.class.create({
      data: {
        id: createId(),
        name,
        year,
        subjectId
      }
    })

    logger.info('Class created', { subjectId, name })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to create class', { error, subjectId })
    return handleServerActionError(error)
  }
})

/**
 * Update a class (HOD/Admin only)
 */
export const updateClass = withRole(['admin', 'hod'], async (
  classId: string,
  subjectId: string,
  formData: FormData
) => {
  try {
    await checkSubjectAccess(subjectId)

    const name = formData.get('name') as string
    const year = formData.get('year') as string

    await prisma.class.update({
      where: { id: classId },
      data: { name, year }
    })

    logger.info('Class updated', { classId })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to update class', { error, classId })
    return handleServerActionError(error)
  }
})

/**
 * Delete a class (HOD/Admin only)
 */
export const deleteClass = withRole(['admin', 'hod'], async (
  classId: string,
  subjectId: string
) => {
  try {
    await checkSubjectAccess(subjectId)

    await prisma.class.delete({
      where: { id: classId }
    })

    logger.info('Class deleted', { classId })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete class', { error, classId })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Comment Group Management Actions
// ============================================================================

/**
 * Create a comment group (HOD/Admin only)
 */
export const createCommentGroup = withRole(['admin', 'hod'], async (
  subjectId: string,
  formData: FormData
) => {
  try {
    await checkSubjectAccess(subjectId)

    const data = {
      subjectId,
      name: formData.get('name') as string,
      title: formData.get('title') as string
    }

    const validation = validateFormData(CreateCommentGroupSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    // Get the current max displayOrder
    const maxOrder = await prisma.commentGroup.aggregate({
      where: { subjectId },
      _max: { displayOrder: true }
    })

    await prisma.commentGroup.create({
      data: {
        id: createId(),
        name: validated.name,
        title: validated.title,
        subjectId: validated.subjectId,
        displayOrder: (maxOrder._max.displayOrder || 0) + 1
      }
    })

    logger.info('Comment group created', { subjectId, name: validated.name })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to create comment group', { error, subjectId })
    return handleServerActionError(error)
  }
})

/**
 * Update a comment group (HOD/Admin only)
 */
export const updateCommentGroup = withRole(['admin', 'hod'], async (
  groupId: string,
  subjectId: string,
  formData: FormData
) => {
  try {
    await checkSubjectAccess(subjectId)

    const data = {
      groupId,
      name: formData.get('name') as string,
      title: formData.get('title') as string
    }

    const validation = validateFormData(UpdateCommentGroupSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    await prisma.commentGroup.update({
      where: { id: groupId },
      data: {
        name: validated.name,
        title: validated.title
      }
    })

    logger.info('Comment group updated', { groupId })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to update comment group', { error, groupId })
    return handleServerActionError(error)
  }
})

/**
 * Delete a comment group (HOD/Admin only)
 */
export const deleteCommentGroup = withRole(['admin', 'hod'], async (
  groupId: string,
  subjectId: string
) => {
  try {
    await checkSubjectAccess(subjectId)

    const validation = validateFormData(DeleteCommentGroupSchema, { groupId })
    if (!validation.success) {
      return validation
    }

    await prisma.commentGroup.delete({
      where: { id: groupId }
    })

    logger.info('Comment group deleted', { groupId })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete comment group', { error, groupId })
    return handleServerActionError(error)
  }
})

/**
 * Reorder comment groups (HOD/Admin only)
 */
export const reorderCommentGroups = withRole(['admin', 'hod'], async (
  subjectId: string,
  items: { id: string; order: number }[]
) => {
  try {
    await checkSubjectAccess(subjectId)

    const validation = validateFormData(ReorderCommentGroupsSchema, {
      subjectId,
      groupIds: items.map(i => i.id)
    })
    if (!validation.success) {
      return validation
    }

    // Update display order for each group
    await Promise.all(
      items.map(item =>
        prisma.commentGroup.update({
          where: { id: item.id },
          data: { displayOrder: item.order }
        })
      )
    )

    logger.info('Comment groups reordered', { subjectId, count: items.length })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to reorder comment groups', { error, subjectId })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Comment Option Management Actions
// ============================================================================

/**
 * Create a comment option (HOD/Admin only)
 */
export const createComment = withRole(['admin', 'hod'], async (
  groupId: string,
  subjectId: string,
  formData: FormData
) => {
  try {
    await checkSubjectAccess(subjectId)

    const data = {
      groupId,
      code: formData.get('code') as string,
      text: formData.get('text') as string
    }

    const validation = validateFormData(CreateCommentOptionSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    // Get the current max displayOrder
    const maxOrder = await prisma.commentOption.aggregate({
      where: { groupId },
      _max: { displayOrder: true }
    })

    await prisma.commentOption.create({
      data: {
        id: createId(),
        code: validated.code,
        text: validated.text,
        groupId: validated.groupId,
        displayOrder: (maxOrder._max.displayOrder || 0) + 1
      }
    })

    logger.info('Comment option created', { groupId, code: validated.code })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to create comment option', { error, groupId })
    return handleServerActionError(error)
  }
})

/**
 * Update a comment option (HOD/Admin only)
 */
export const updateComment = withRole(['admin', 'hod'], async (
  commentId: string,
  subjectId: string,
  groupId: string,
  formData: FormData
) => {
  try {
    await checkSubjectAccess(subjectId)

    const data = {
      optionId: commentId,
      code: formData.get('code') as string,
      text: formData.get('text') as string
    }

    const validation = validateFormData(UpdateCommentOptionSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    await prisma.commentOption.update({
      where: { id: commentId },
      data: {
        code: validated.code,
        text: validated.text
      }
    })

    logger.info('Comment option updated', { commentId })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to update comment option', { error, commentId })
    return handleServerActionError(error)
  }
})

/**
 * Delete a comment option (HOD/Admin only)
 */
export const deleteComment = withRole(['admin', 'hod'], async (
  commentId: string,
  subjectId: string,
  groupId: string
) => {
  try {
    await checkSubjectAccess(subjectId)

    const validation = validateFormData(DeleteCommentOptionSchema, { optionId: commentId })
    if (!validation.success) {
      return validation
    }

    await prisma.commentOption.delete({
      where: { id: commentId }
    })

    logger.info('Comment option deleted', { commentId })
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete comment option', { error, commentId })
    return handleServerActionError(error)
  }
})

/**
 * Reorder comment options (HOD/Admin only)
 */
export const reorderComments = withRole(['admin', 'hod'], async (
  groupId: string,
  items: { id: string; order: number }[]
) => {
  try {
    const validation = validateFormData(ReorderCommentOptionsSchema, {
      groupId,
      optionIds: items.map(i => i.id)
    })
    if (!validation.success) {
      return validation
    }

    // Update display order for each option
    await Promise.all(
      items.map(item =>
        prisma.commentOption.update({
          where: { id: item.id },
          data: { displayOrder: item.order }
        })
      )
    )

    logger.info('Comment options reordered', { groupId, count: items.length })

    // Get subject ID for revalidation
    const group = await prisma.commentGroup.findUnique({
      where: { id: groupId },
      select: { subjectId: true }
    })
    if (group) {
      revalidatePath(`/hod/subject/${group.subjectId}`)
    }

    return { success: true }
  } catch (error) {
    logger.error('Failed to reorder comment options', { error, groupId })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Assignment Management Actions
// ============================================================================

/**
 * Create an assignment (HOD/Admin only)
 */
export const createAssignment = withRole(['admin', 'hod'], async (
  classId: string,
  formData: FormData
) => {
  try {
    const pupilId = formData.get('pupilId') as string
    const eoyLevel = formData.get('eoyLevel') as string
    const targetLevel = formData.get('targetLevel') as string

    if (!pupilId) {
      return {
        success: false,
        error: 'Pupil ID is required',
        code: 'VALIDATION_ERROR'
      }
    }

    await prisma.assignment.create({
      data: {
        id: createId(),
        pupilId,
        classId,
        eoyLevel,
        targetLevel
      }
    })

    logger.info('Assignment created', { classId, pupilId })
    revalidatePath(`/class/${classId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to create assignment', { error, classId })
    return handleServerActionError(error)
  }
})

/**
 * Update an assignment (HOD/Admin only)
 */
export const updateAssignment = withRole(['admin', 'hod'], async (
  assignmentId: string,
  classId: string,
  formData: FormData
) => {
  try {
    const eoyLevel = formData.get('eoyLevel') as string
    const targetLevel = formData.get('targetLevel') as string
    const actualLevel = formData.get('actualLevel') as string

    await prisma.assignment.update({
      where: { id: assignmentId },
      data: {
        eoyLevel,
        targetLevel,
        actualLevel
      }
    })

    logger.info('Assignment updated', { assignmentId })
    revalidatePath(`/class/${classId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to update assignment', { error, assignmentId })
    return handleServerActionError(error)
  }
})

/**
 * Delete an assignment (HOD/Admin only)
 */
export const deleteAssignment = withRole(['admin', 'hod'], async (
  assignmentId: string,
  classId: string
) => {
  try {
    await prisma.assignment.delete({
      where: { id: assignmentId }
    })

    logger.info('Assignment deleted', { assignmentId })
    revalidatePath(`/class/${classId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete assignment', { error, assignmentId })
    return handleServerActionError(error)
  }
})

// Export aliases for backward compatibility
export {
  createAssignment as createStudent,
  updateAssignment as updateStudent,
  deleteAssignment as deleteStudent
}
