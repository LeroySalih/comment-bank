'use server'

import { prisma } from '@/lib/prisma'
import { revalidatePath } from 'next/cache'
import { createId } from '@paralleldrive/cuid2'

export async function updateAssignmentCode(
  assignmentId: string,
  groupId: string,
  code: string | null
) {
  try {
    if (code === null || code === '') {
      await (prisma as any).pupilCode.delete({
        where: { assignmentId_groupId: { assignmentId, groupId } }
      }).catch(() => {});
    } else {
      await (prisma as any).pupilCode.upsert({
        where: { assignmentId_groupId: { assignmentId, groupId } },
        update: { code },
        create: { id: createId(), assignmentId, groupId, code }
      })
    }
    
    revalidatePath(`/student/${assignmentId}`)
    revalidatePath('/')
    
    // Also revalidate the class page
    const assignment = await (prisma as any).assignment.findUnique({
      where: { id: assignmentId },
      select: { classId: true }
    });
    if (assignment?.classId) {
      revalidatePath(`/class/${assignment.classId}`);
    }

    return { success: true }
  } catch (error) {
    console.error('Failed to update code:', error)
    return { success: false, error: 'Database error' }
  }
}

export async function updateAssignmentCommentText(assignmentId: string, comment: string) {
  try {
    // Get current assignment to check status
    const currentAssignment = await prisma.assignment.findUnique({
      where: { id: assignmentId },
      select: { checkStatus: true, classId: true }
    })

    if (!currentAssignment) {
      return { success: false, error: 'Assignment not found' }
    }

    // Determine if status needs to change to required_check
    // If teacher is editing the comment (saving to finalComment), it needs review
    // unless it's already been reviewed and approved
    const currentStatus = currentAssignment.checkStatus || 'not_required'
    let newStatus = currentStatus

    // If status is "not_required", change to "required_check" since teacher is now editing
    // If status is "checked_rejected", also change to "required_check" (resubmitting)
    // If status is "checked_ok" and teacher edits again, it needs re-review
    if (currentStatus === 'not_required' || currentStatus === 'checked_rejected' || currentStatus === 'checked_ok') {
      newStatus = 'required_check'
    }

    const updated = await prisma.assignment.update({
      where: { id: assignmentId },
      data: {
        finalComment: comment,
        checkStatus: newStatus,
        // Clear check data if we're changing status to required_check
        ...(newStatus === 'required_check' && currentStatus !== 'required_check' ? {
          checkNote: null,
          checkedAt: null,
          checkedById: null
        } : {})
      }
    })

    // Don't call revalidatePath here - it causes stale data to be fetched
    // The client has the correct data in local state
    return { success: true, finalComment: updated.finalComment }
  } catch (error) {
    console.error('Failed to update comment text:', error)
    return { success: false, error: error instanceof Error ? error.message : 'Database error' }
  }
}

export async function revertAssignmentComment(assignmentId: string) {
  try {
    // Get current assignment to verify it exists
    const currentAssignment = await prisma.assignment.findUnique({
      where: { id: assignmentId },
      select: { classId: true }
    });

    if (!currentAssignment) {
      return { success: false, error: 'Assignment not found' }
    }

    // Clear the finalComment and reset status to not_required
    await prisma.assignment.update({
      where: { id: assignmentId },
      data: {
        finalComment: null,
        checkStatus: 'not_required',
        checkNote: null,
        checkedAt: null,
        checkedById: null
      }
    });

    return { success: true }
  } catch (error) {
    console.error('Failed to revert comment:', error)
    return { success: false, error: error instanceof Error ? error.message : 'Database error' }
  }
}
