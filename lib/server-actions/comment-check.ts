"use server"

import { revalidatePath } from 'next/cache'
import { pool } from '@/lib/db'
import { withRole } from '@/lib/auth/with-role'
import { handleServerActionError, ForbiddenError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import { getServerSession } from 'next-auth'
import { authOptions } from '@/app/api/auth/[...nextauth]/route'
import { createAuditLog } from '@/lib/audit-log'

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
  const { rows } = await pool.query(
    `SELECT 1 FROM "Subject" s
     JOIN "_SubjectToUser" su ON su."A" = s.id
     WHERE s.id = $1 AND su."B" = $2`,
    [subjectId, session.user.id]
  )

  if (rows.length === 0) {
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

    const { rows: aRows } = await pool.query(
      `SELECT a.*, c.id as "class_id", c.name as "class_name", c."subjectId" as "class_subjectId",
              s.id as "subject_id", s.code as "subject_code"
       FROM "Assignment" a
       JOIN "Class" c ON c.id = a."classId"
       LEFT JOIN "Subject" s ON s.id = c."subjectId"
       WHERE a.id = $1`,
      [assignmentId]
    )

    if (aRows.length === 0) {
      return { success: false, error: 'Assignment not found', code: 'NOT_FOUND' }
    }

    const row = aRows[0]
    const assignment = {
      ...row,
      classId: row.classId,
      checkStatus: row.checkStatus,
      Class: {
        id: row.class_id,
        name: row.class_name,
        subjectId: row.class_subjectId,
        Subject: row.subject_id ? { id: row.subject_id, code: row.subject_code } : null
      }
    }

    // Check HoD has access to this subject
    await checkSubjectAccess(assignment.Class.subjectId)

    // Validation
    if (status === 'checked_rejected' && !note?.trim()) {
      return { success: false, error: 'A note is required when rejecting a comment', code: 'VALIDATION_ERROR' }
    }

    await pool.query(
      `UPDATE "Assignment"
       SET "checkStatus" = $1, "checkNote" = $2, "checkedAt" = $3, "checkedById" = $4
       WHERE id = $5`,
      [status, note?.trim() || null, new Date(), session?.user?.id ?? null, assignmentId]
    )

    // Audit log
    await createAuditLog({
      action: status === 'checked_ok' ? 'review_comment_approved' : 'review_comment_rejected',
      entityType: 'assignment',
      entityId: assignmentId,
      details: {
        className: assignment.Class.name,
        subjectCode: assignment.Class.Subject?.code,
        previousStatus: assignment.checkStatus,
        newStatus: status,
        note: note?.trim() || null
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

    const { rows: aRows } = await pool.query(
      `SELECT a.*, c.id as "class_id", c.name as "class_name", c."subjectId" as "class_subjectId"
       FROM "Assignment" a
       JOIN "Class" c ON c.id = a."classId"
       WHERE a.id = $1`,
      [assignmentId]
    )

    if (aRows.length === 0) {
      return { success: false, error: 'Assignment not found', code: 'NOT_FOUND' }
    }

    const row = aRows[0]
    const assignment = {
      ...row,
      classId: row.classId,
      checkStatus: row.checkStatus,
      Class: {
        id: row.class_id,
        name: row.class_name,
        subjectId: row.class_subjectId
      }
    }

    // Only allow reset if currently rejected
    if (assignment.checkStatus !== 'checked_rejected') {
      return { success: false, error: 'Can only resubmit rejected comments', code: 'INVALID_STATUS' }
    }

    await pool.query(
      `UPDATE "Assignment"
       SET "checkStatus" = 'required_check', "checkNote" = NULL, "checkedAt" = NULL, "checkedById" = NULL
       WHERE id = $1`,
      [assignmentId]
    )

    // Audit log
    await createAuditLog({
      action: 'reset_comment_status',
      entityType: 'assignment',
      entityId: assignmentId,
      details: {
        className: assignment.Class.name,
        previousStatus: assignment.checkStatus,
        newStatus: 'required_check'
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
  const checkStatusValue = statusFilter ?? 'required_check'

  // Fetch assignments joined with Pupil and Class
  const { rows: assignmentRows } = await pool.query(
    `SELECT a.*,
            p."admissionNumber" as "pupil_admissionNumber", p."firstName" as "pupil_firstName",
            p."lastName" as "pupil_lastName", p.gender as "pupil_gender", p."isActive" as "pupil_isActive", p.form as "pupil_form",
            c.id as "class_id", c.name as "class_name", c.year as "class_year", c."subjectId" as "class_subjectId",
            u.id as "checker_id", u.username as "checker_username"
     FROM "Assignment" a
     JOIN "Class" c ON c.id = a."classId"
     JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
     LEFT JOIN "User" u ON u.id = a."checkedById"
     WHERE c."subjectId" = $1 AND a."checkStatus" = $2
     ORDER BY a."checkStatus" ASC, a."checkedAt" DESC`,
    [subjectId, checkStatusValue]
  )

  if (assignmentRows.length === 0) {
    return []
  }

  const assignmentIds = assignmentRows.map(r => r.id)

  // Fetch PupilCodes with CommentGroups and CommentOptions
  const { rows: pupilCodeRows } = await pool.query(
    `SELECT pc.*, cg.id as "cg_id", cg.name as "cg_name", cg.title as "cg_title",
            cg."displayOrder" as "cg_displayOrder", cg."subjectId" as "cg_subjectId",
            cg."isLinked" as "cg_isLinked", cg."linkedField" as "cg_linkedField",
            co.id as "co_id", co.code as "co_code", co.text as "co_text",
            co."displayOrder" as "co_displayOrder", co."groupId" as "co_groupId"
     FROM "PupilCode" pc
     JOIN "CommentGroup" cg ON cg.id = pc."groupId"
     LEFT JOIN "CommentOption" co ON co."groupId" = pc."groupId"
     WHERE pc."assignmentId" = ANY($1::text[])`,
    [assignmentIds]
  )

  // Group PupilCodes by assignmentId, then build CommentGroup with CommentOption array
  const pupilCodeMap = new Map<string, any[]>()
  for (const row of pupilCodeRows) {
    const existing = pupilCodeMap.get(row.assignmentId)
    // Check if this groupId already exists in the array
    if (!existing) {
      pupilCodeMap.set(row.assignmentId, [])
    }
    const arr = pupilCodeMap.get(row.assignmentId)!
    let pcEntry = arr.find(e => e.id === row.id)
    if (!pcEntry) {
      pcEntry = {
        id: row.id,
        assignmentId: row.assignmentId,
        groupId: row.groupId,
        code: row.code,
        CommentGroup: {
          id: row.cg_id,
          name: row.cg_name,
          title: row.cg_title,
          displayOrder: row.cg_displayOrder,
          subjectId: row.cg_subjectId,
          isLinked: row.cg_isLinked,
          linkedField: row.cg_linkedField,
          CommentOption: [] as any[]
        }
      }
      arr.push(pcEntry)
    }
    if (row.co_id) {
      pcEntry.CommentGroup.CommentOption.push({
        id: row.co_id,
        code: row.co_code,
        text: row.co_text,
        displayOrder: row.co_displayOrder,
        groupId: row.co_groupId
      })
    }
  }

  // Assemble final result
  return assignmentRows.map(row => ({
    id: row.id,
    pupilId: row.pupilId,
    classId: row.classId,
    eoyLevel: row.eoyLevel,
    targetLevel: row.targetLevel,
    actualLevel: row.actualLevel,
    finalComment: row.finalComment,
    linkedData: row.linkedData,
    checkStatus: row.checkStatus,
    checkNote: row.checkNote,
    checkedAt: row.checkedAt,
    checkedById: row.checkedById,
    Pupil: {
      admissionNumber: row.pupil_admissionNumber,
      firstName: row.pupil_firstName,
      lastName: row.pupil_lastName,
      gender: row.pupil_gender,
      isActive: row.pupil_isActive,
      form: row.pupil_form
    },
    Class: {
      id: row.class_id,
      name: row.class_name,
      year: row.class_year,
      subjectId: row.class_subjectId
    },
    PupilCode: pupilCodeMap.get(row.id) ?? [],
    CheckedBy: row.checker_id ? { id: row.checker_id, username: row.checker_username } : null
  }))
}

/**
 * Get review statistics for a subject
 */
export async function getReviewStats(subjectId: string) {
  const { rows } = await pool.query(
    `SELECT a."checkStatus", COUNT(*) as count
     FROM "Assignment" a
     JOIN "Class" c ON c.id = a."classId"
     WHERE c."subjectId" = $1
     GROUP BY a."checkStatus"`,
    [subjectId]
  )

  const find = (status: string) => {
    const r = rows.find(s => s.checkStatus === status)
    return r ? Number(r.count) : 0
  }

  return {
    notRequired: find('not_required'),
    pendingReview: find('required_check'),
    approved: find('checked_ok'),
    rejected: find('checked_rejected'),
    total: rows.reduce((sum, s) => sum + Number(s.count), 0)
  }
}
