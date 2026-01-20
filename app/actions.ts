'use server'

import { prisma } from '@/lib/prisma'
import { revalidatePath } from 'next/cache'

export async function updateStudentComment(
  studentId: string, 
  groupName: string, 
  code: string | null
) {
  const data: Record<string, string | null> = {}

  // Map Group Name to DB Column
  // Note: This relies on the hardcoded columns in the schema
  switch (groupName) {
    case 'WP':
      data.wpCode = code
      break
    case 'TH':
      data.thCode = code
      break
    case 'PS':
      data.psCode = code
      break
    case 'OA':
      data.oaCode = code
      break
    default:
      // For groups not in the schema columns, we can't save them in this iteration 
      // without schema changes. Ignoring for now as per legacy support.
      return { success: false, error: 'Unknown group' }
  }

  try {
    await prisma.student.update({
      where: { id: studentId },
      data
    })
    
    revalidatePath(`/student/${studentId}`)
    revalidatePath('/') // Revalidate dashboard for completion stats
    return { success: true }
  } catch (error) {
    console.error('Failed to update comment:', error)
    return { success: false, error: 'Database error' }
  }
}

export async function updateStudentCommentText(studentId: string, comment: string) {
  try {
    await prisma.student.update({
      where: { id: studentId },
      data: { comment }
    })
    
    revalidatePath(`/student/${studentId}`)
    return { success: true }
  } catch (error) {
    console.error('Failed to update comment text:', error)
    return { success: false, error: 'Database error' }
  }
}
