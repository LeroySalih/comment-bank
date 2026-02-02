"use server"

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { withRole } from '@/lib/auth/with-role'
import { handleServerActionError, ForbiddenError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import { getServerSession } from 'next-auth'
import { authOptions } from '@/app/api/auth/[...nextauth]/route'

// Comment check status types
export type CheckStatus = 'not_required' | 'required_check' | 'checked_ok' | 'checked_rejected'

/**
 * Helper to check if user has HOD access to a subject
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
    throw new ForbiddenError('You do not have access to review comments for this subject')
  }
}

// ============================================================================
// Comment Review Actions
// ============================================================================

/**
 * HoD reviews a comment and sets it to checked_ok or checked_rejected
 */
export const reviewComment = withRole(['admin', 'hod'], async (
  assignmentId: string,
  status: 'checked_ok' | 'checked_rejected',
  note?: string
) => {
  try {
    const session = await getServerSession(authOptions)

    const assignment = await prisma.assignment.findUnique({
      where: { id: assignmentId },
      include: { Class: { include: { Subject: true } } }
    })

    if (!assignment) {
      return { success: false, error: 'Assignment not found', code: 'NOT_FOUND' }
    }

    // Check HoD has access to this subject
    await checkSubjectAccess(assignment.Class.subjectId)

    // Validation
    if (status === 'checked_rejected' && !note?.trim()) {
      return { success: false, error: 'A note is required when rejecting a comment', code: 'VALIDATION_ERROR' }
    }

    await prisma.assignment.update({
      where: { id: assignmentId },
      data: {
        checkStatus: status,
        checkNote: note?.trim() || null,
        checkedAt: new Date(),
        checkedById: session?.user?.id
      }
    })

    logger.info('Comment reviewed', { assignmentId, status, reviewerId: session?.user?.id })

    revalidatePath(`/student/${assignmentId}`)
    revalidatePath(`/class/${assignment.classId}`)
    revalidatePath(`/hod/subject/${assignment.Class.subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to review comment', { error, assignmentId })
    return handleServerActionError(error)
  }
})

/**
 * Reset a comment status (e.g., after teacher fixes a rejected comment)
 * Teachers can reset their own rejected comments, HoDs can reset any
 */
export const resetCommentStatus = withRole(['admin', 'hod', 'teacher'], async (
  assignmentId: string
) => {
  try {
    const session = await getServerSession(authOptions)

    const assignment = await prisma.assignment.findUnique({
      where: { id: assignmentId },
      include: { Class: true }
    })

    if (!assignment) {
      return { success: false, error: 'Assignment not found', code: 'NOT_FOUND' }
    }

    // Only allow reset if currently rejected
    if (assignment.checkStatus !== 'checked_rejected') {
      return { success: false, error: 'Can only resubmit rejected comments', code: 'INVALID_STATUS' }
    }

    await prisma.assignment.update({
      where: { id: assignmentId },
      data: {
        checkStatus: 'required_check',
        checkNote: null,
        checkedAt: null,
        checkedById: null
      }
    })

    logger.info('Comment status reset', { assignmentId, userId: session?.user?.id })

    revalidatePath(`/student/${assignmentId}`)
    revalidatePath(`/class/${assignment.classId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to reset comment status', { error, assignmentId })
    return handleServerActionError(error)
  }
})

/**
 * Get assignments requiring review for a subject
 */
export async function getAssignmentsForReview(subjectId: string, statusFilter?: CheckStatus) {
  const whereClause: any = {
    Class: { subjectId }
  }

  if (statusFilter) {
    whereClause.checkStatus = statusFilter
  } else {
    // Default: show only those needing review
    whereClause.checkStatus = 'required_check'
  }

  const assignments = await prisma.assignment.findMany({
    where: whereClause,
    include: {
      Pupil: true,
      Class: true,
      PupilCode: {
        include: {
          CommentGroup: {
            include: { CommentOption: true }
          }
        }
      },
      CheckedBy: {
        select: { id: true, username: true }
      }
    },
    orderBy: [
      { checkStatus: 'asc' },
      { checkedAt: 'desc' }
    ]
  })

  return assignments
}

/**
 * Get review statistics for a subject
 */
export async function getReviewStats(subjectId: string) {
  const stats = await prisma.assignment.groupBy({
    by: ['checkStatus'],
    where: {
      Class: { subjectId }
    },
    _count: true
  })

  return {
    notRequired: stats.find(s => s.checkStatus === 'not_required')?._count || 0,
    pendingReview: stats.find(s => s.checkStatus === 'required_check')?._count || 0,
    approved: stats.find(s => s.checkStatus === 'checked_ok')?._count || 0,
    rejected: stats.find(s => s.checkStatus === 'checked_rejected')?._count || 0,
    total: stats.reduce((sum, s) => sum + s._count, 0)
  }
}
