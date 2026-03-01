"use server"

import { revalidatePath } from 'next/cache'
import { pool } from '@/lib/db'
import { withRole } from '@/lib/auth/with-role'
import { handleServerActionError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import { createAuditLog } from '@/lib/audit-log'
import * as XLSX from 'xlsx'

/**
 * Get distinct linked field names from Assignment.linkedData across all assignments.
 * Used to populate the admin dropdown when creating linked groups.
 */
export const getAvailableLinkedFields = withRole('admin', async () => {
  try {
    const { rows } = await pool.query<{ field_name: string }>(
      `SELECT DISTINCT jsonb_object_keys("linkedData") as field_name
       FROM "Assignment" WHERE "linkedData" IS NOT NULL`
    )

    return {
      success: true,
      fields: rows.map(r => r.field_name).sort()
    }
  } catch (error) {
    logger.error('Failed to get available linked fields', { error })
    return { success: true, fields: [] as string[] }
  }
})

/**
 * Upload linked data for a class from an Excel/CSV file.
 * Expects columns: AdmissionNumber + arbitrary data columns (Behaviour, Effort, Homework, etc.)
 * Updates assignment.linkedData for matching students.
 */
export const uploadClassLinkedData = withRole('admin', async (
  classId: string,
  base64Content: string
) => {
  try {
    if (!classId) {
      return { success: false, error: 'Class ID is required' }
    }

    const workbook = XLSX.read(base64Content, { type: 'base64' })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json<Record<string, any>>(sheet)

    if (rows.length === 0) {
      return { success: false, error: 'No data found in the uploaded file' }
    }

    // Get all assignments for this class
    const { rows: assignments } = await pool.query<{ id: string; pupilId: string }>(
      `SELECT id, "pupilId" FROM "Assignment" WHERE "classId" = $1`,
      [classId]
    )

    const assignmentByPupilId = new Map(
      assignments.map(a => [a.pupilId, a.id])
    )

    const PII_COLUMNS = new Set([
      'admissionnumber', 'admno', 'name', 'firstname', 'lastname',
      'surname', 'forename', 'email', 'dob', 'dateofbirth',
      'address', 'phone', 'telephone', 'mobile', 'postcode', 'zipcode'
    ])

    let matchedCount = 0
    const unmatchedRows: string[] = []
    const detectedFields = new Set<string>()
    const skippedPiiColumns = new Set<string>()

    for (const row of rows) {
      const admissionNumber = String(
        row.AdmissionNumber || row.admissionNumber || row.AdmNo || row.admNo || ''
      ).trim()

      if (!admissionNumber) {
        unmatchedRows.push('Row with no admission number')
        continue
      }

      const assignmentId = assignmentByPupilId.get(admissionNumber)
      if (!assignmentId) {
        unmatchedRows.push(admissionNumber)
        continue
      }

      // Build linkedData from remaining columns, lowercasing keys
      const linkedData: Record<string, string> = {}
      for (const [key, value] of Object.entries(row)) {
        const lowerKey = key.toLowerCase()
        if (PII_COLUMNS.has(lowerKey)) {
          skippedPiiColumns.add(lowerKey)
          continue
        }
        linkedData[lowerKey] = String(value).trim()
        detectedFields.add(lowerKey)
      }

      await pool.query(
        `UPDATE "Assignment" SET "linkedData" = $1 WHERE id = $2`,
        [JSON.stringify(linkedData), assignmentId]
      )

      matchedCount++
    }

    await createAuditLog({
      action: 'upload_linked_data',
      entityType: 'assignment',
      details: {
        classId,
        action: 'upload_linked_data',
        matchedCount,
        unmatchedCount: unmatchedRows.length,
        detectedFields: Array.from(detectedFields),
        skippedPiiColumns: Array.from(skippedPiiColumns)
      }
    })

    logger.info('Linked data uploaded', { classId, matchedCount })
    revalidatePath(`/class/${classId}`)
    revalidatePath('/admin')

    return {
      success: true,
      matchedCount,
      unmatchedRows,
      detectedFields: Array.from(detectedFields).sort()
    }
  } catch (error) {
    logger.error('Failed to upload linked data', { error, classId })
    return handleServerActionError(error)
  }
})
