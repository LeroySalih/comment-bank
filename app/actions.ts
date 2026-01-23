'use server'

import { prisma } from '@/lib/prisma'
import { revalidatePath } from 'next/cache'

export async function updateAssignmentCode(
  assignmentId: string, 
  groupId: string, 
  code: string | null
) {
  try {
    if (code === null) {
      await (prisma as any).pupilCode.delete({
        where: { assignmentId_groupId: { assignmentId, groupId } }
      }).catch(() => {});
    } else {
      await (prisma as any).pupilCode.upsert({
        where: { assignmentId_groupId: { assignmentId, groupId } },
        update: { code },
        create: { assignmentId, groupId, code }
      })
    }
    
    revalidatePath(`/student/${assignmentId}`)
    revalidatePath('/')
    return { success: true }
  } catch (error) {
    console.error('Failed to update code:', error)
    return { success: false, error: 'Database error' }
  }
}

export async function updateAssignmentCommentText(assignmentId: string, comment: string) {
  try {
    await (prisma as any).assignment.update({
      where: { id: assignmentId },
      data: { finalComment: comment }
    })
    
    revalidatePath(`/student/${assignmentId}`)
    return { success: true }
  } catch (error) {
    console.error('Failed to update comment text:', error)
    return { success: false, error: 'Database error' }
  }
}
