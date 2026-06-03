'use server'

import { pool } from '@/lib/db'
import { revalidatePath } from 'next/cache'
import { createId } from '@paralleldrive/cuid2'
import { createAuditLog, logDataChange } from '@/lib/audit-log'
import { encrypt } from '@/lib/encryption'

export async function updateAssignmentCode(
  assignmentId: string,
  groupId: string,
  code: string | null
) {
  try {
    // Guard: reject if group is linked
    const { rows: groupRows } = await pool.query(
      `SELECT id, "isLinked" FROM "CommentGroup" WHERE id = $1`,
      [groupId]
    )
    const group = groupRows[0] ?? null
    if (group && group.isLinked) {
      return { success: false, error: 'Cannot manually update a linked group code' }
    }

    // Get current code for audit log
    const { rows: currentRows } = await pool.query(
      `SELECT code FROM "PupilCode" WHERE "assignmentId" = $1 AND "groupId" = $2`,
      [assignmentId, groupId]
    )
    const oldCode = currentRows[0]?.code ?? null

    if (code === null || code === '') {
      await pool.query(
        `DELETE FROM "PupilCode" WHERE "assignmentId" = $1 AND "groupId" = $2`,
        [assignmentId, groupId]
      )
    } else {
      await pool.query(
        `INSERT INTO "PupilCode" (id, "assignmentId", "groupId", code)
         VALUES ($1, $2, $3, $4)
         ON CONFLICT ("assignmentId", "groupId") DO UPDATE SET code = EXCLUDED.code`,
        [createId(), assignmentId, groupId, code]
      )
    }

    // Audit log - only log if code actually changed
    if (oldCode !== code) {
      await logDataChange(
        'update_assignment_code',
        'assignment',
        assignmentId,
        { groupId, code: oldCode },
        { groupId, code }
      )
    }

    revalidatePath(`/student/${assignmentId}`)
    revalidatePath('/')

    // Also revalidate the class page
    const { rows: assignmentRows } = await pool.query(
      `SELECT "classId" FROM "Assignment" WHERE id = $1`,
      [assignmentId]
    )
    if (assignmentRows[0]?.classId) {
      revalidatePath(`/class/${assignmentRows[0].classId}`)
    }

    return { success: true }
  } catch (error) {
    console.error('Failed to update code:', error)
    return { success: false, error: 'Database error' }
  }
}

export async function updateCommonAssignmentCode(
  assignmentId: string,
  commonGroupId: string,
  code: string | null
) {
  try {
    // Guard: reject if group is linked
    const { rows: cgRows } = await pool.query(
      `SELECT id, "isLinked" FROM "CommonCommentGroup" WHERE id = $1`,
      [commonGroupId]
    )
    const commonGroup = cgRows[0] ?? null
    if (commonGroup && commonGroup.isLinked) {
      return { success: false, error: 'Cannot manually update a linked group code' }
    }

    const { rows: currentRows } = await pool.query(
      `SELECT code FROM "CommonPupilCode" WHERE "assignmentId" = $1 AND "commonGroupId" = $2`,
      [assignmentId, commonGroupId]
    )
    const oldCode = currentRows[0]?.code ?? null

    if (code === null || code === '') {
      await pool.query(
        `DELETE FROM "CommonPupilCode" WHERE "assignmentId" = $1 AND "commonGroupId" = $2`,
        [assignmentId, commonGroupId]
      )
    } else {
      await pool.query(
        `INSERT INTO "CommonPupilCode" (id, "assignmentId", "commonGroupId", code)
         VALUES ($1, $2, $3, $4)
         ON CONFLICT ("assignmentId", "commonGroupId") DO UPDATE SET code = EXCLUDED.code`,
        [createId(), assignmentId, commonGroupId, code]
      )
    }

    if (oldCode !== code) {
      await logDataChange(
        'update_common_assignment_code',
        'assignment',
        assignmentId,
        { commonGroupId, code: oldCode },
        { commonGroupId, code }
      )
    }

    revalidatePath(`/student/${assignmentId}`)
    revalidatePath('/')

    const { rows: assignmentRows } = await pool.query(
      `SELECT "classId" FROM "Assignment" WHERE id = $1`,
      [assignmentId]
    )
    if (assignmentRows[0]?.classId) {
      revalidatePath(`/class/${assignmentRows[0].classId}`)
    }

    return { success: true }
  } catch (error) {
    console.error('Failed to update common code:', error)
    return { success: false, error: 'Database error' }
  }
}

export async function updateAssignmentCommentText(assignmentId: string, comment: string, targetStatus?: 'required_check' | 'not_required') {
  try {
    // Get current assignment to check status
    const { rows: currentRows } = await pool.query(
      `SELECT "checkStatus", "classId" FROM "Assignment" WHERE id = $1`,
      [assignmentId]
    )

    if (currentRows.length === 0) {
      return { success: false, error: 'Assignment not found' }
    }

    const currentAssignment = currentRows[0]

    // Determine if status needs to change
    const currentStatus = currentAssignment.checkStatus || 'not_required'
    let newStatus: string

    if (targetStatus) {
      newStatus = targetStatus
    } else if (currentStatus === 'not_required' || currentStatus === 'checked_rejected' || currentStatus === 'checked_ok') {
      newStatus = 'required_check'
    } else {
      newStatus = currentStatus
    }

    const clearCheckData = newStatus === 'required_check' && currentStatus !== 'required_check'

    const { rows: updatedRows } = await pool.query(
      `UPDATE "Assignment"
       SET "finalComment" = $1,
           "checkStatus" = $2
           ${clearCheckData ? ', "checkNote" = NULL, "checkedAt" = NULL, "checkedById" = NULL' : ''}
       WHERE id = $3
       RETURNING "finalComment"`,
      [encrypt(comment), newStatus, assignmentId]
    )

    // Audit log
    await logDataChange(
      'update_assignment_comment',
      'assignment',
      assignmentId,
      { finalComment: currentAssignment.checkStatus === 'not_required' ? null : '(previous comment)', checkStatus: currentStatus },
      { finalComment: '(comment updated)', checkStatus: newStatus }
    )

    return { success: true, finalComment: updatedRows[0]?.finalComment ?? null }
  } catch (error) {
    console.error('Failed to update comment text:', error)
    return { success: false, error: error instanceof Error ? error.message : 'Database error' }
  }
}

export async function revertAssignmentComment(assignmentId: string) {
  try {
    // Verify assignment exists and get classId
    const { rows: currentRows } = await pool.query(
      `SELECT "classId" FROM "Assignment" WHERE id = $1`,
      [assignmentId]
    )

    if (currentRows.length === 0) {
      return { success: false, error: 'Assignment not found' }
    }

    // Clear the finalComment and reset status to not_required
    await pool.query(
      `UPDATE "Assignment"
       SET "finalComment" = NULL,
           "checkStatus" = 'not_required',
           "checkNote" = NULL,
           "checkedAt" = NULL,
           "checkedById" = NULL
       WHERE id = $1`,
      [assignmentId]
    )

    // Audit log
    await createAuditLog({
      action: 'revert_assignment_comment',
      entityType: 'assignment',
      entityId: assignmentId,
      details: { note: 'Comment reverted to auto-generated' }
    })

    return { success: true }
  } catch (error) {
    console.error('Failed to revert comment:', error)
    return { success: false, error: error instanceof Error ? error.message : 'Database error' }
  }
}

export async function bulkSetColumnCode(
  classId: string,
  groupId: string,
  code: string,
  groupType: 'subject' | 'common'
): Promise<{ success: boolean; updatedAssignmentIds: string[]; error?: string }> {
  try {
    const { rows: assignmentRows } = await pool.query<{ id: string }>(
      `SELECT a.id FROM "Assignment" a
       JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
       WHERE a."classId" = $1 AND p."isActive" = true`,
      [classId]
    )
    const allIds = assignmentRows.map(r => r.id)
    if (allIds.length === 0) return { success: true, updatedAssignmentIds: [] }

    let alreadySetIds: string[] = []
    if (groupType === 'subject') {
      const { rows } = await pool.query<{ assignmentId: string }>(
        `SELECT "assignmentId" FROM "PupilCode"
         WHERE "assignmentId" = ANY($1::text[]) AND "groupId" = $2`,
        [allIds, groupId]
      )
      alreadySetIds = rows.map(r => r.assignmentId)
    } else {
      const { rows } = await pool.query<{ assignmentId: string }>(
        `SELECT "assignmentId" FROM "CommonPupilCode"
         WHERE "assignmentId" = ANY($1::text[]) AND "commonGroupId" = $2`,
        [allIds, groupId]
      )
      alreadySetIds = rows.map(r => r.assignmentId)
    }

    const alreadySetSet = new Set(alreadySetIds)
    const toUpdate = allIds.filter(id => !alreadySetSet.has(id))
    if (toUpdate.length === 0) return { success: true, updatedAssignmentIds: [] }

    if (groupType === 'subject') {
      for (const assignmentId of toUpdate) {
        await pool.query(
          `INSERT INTO "PupilCode" (id, "assignmentId", "groupId", code)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT ("assignmentId", "groupId") DO NOTHING`,
          [createId(), assignmentId, groupId, code]
        )
      }
    } else {
      for (const assignmentId of toUpdate) {
        await pool.query(
          `INSERT INTO "CommonPupilCode" (id, "assignmentId", "commonGroupId", code)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT ("assignmentId", "commonGroupId") DO NOTHING`,
          [createId(), assignmentId, groupId, code]
        )
      }
    }

    revalidatePath(`/class/${classId}`)

    return { success: true, updatedAssignmentIds: toUpdate }
  } catch (error) {
    console.error('Failed to bulk set column code:', error)
    return { success: false, updatedAssignmentIds: [], error: 'Database error' }
  }
}
