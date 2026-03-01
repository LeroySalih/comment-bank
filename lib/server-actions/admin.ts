"use server"

import { revalidatePath } from 'next/cache'
import { pool } from '@/lib/db'
import { withRole } from '@/lib/auth/with-role'
import { userRepository } from '@/lib/db/repositories/user-repository'
import { pupilRepository } from '@/lib/db/repositories/pupil-repository'
import { subjectRepository } from '@/lib/db/repositories/subject-repository'
import { classRepository } from '@/lib/db/repositories/class-repository'
import { handleServerActionError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import { createAuditLog, logDataChange } from '@/lib/audit-log'
import {
  UpdateUserRolesSchema,
  UpdateUserActiveStatusSchema,
  UpdatePupilSchema,
  CreateSubjectSchema,
  UpdateSubjectSchema,
  DeleteSubjectSchema,
  AssignUserToSubjectSchema,
  AssignTeachersSchema,
  CreateDeadlineSchema,
  UpdateDeadlineSchema,
  DeleteDeadlineSchema,
  validateFormData
} from '@/lib/validation-schemas'
import { createId } from '@paralleldrive/cuid2'
import * as XLSX from 'xlsx'

/**
 * Update user roles (Admin only)
 */
export const updateUserRoles = withRole('admin', async (
  userId: string,
  roleNames: string[]
) => {
  try {
    // Validate input
    const validation = validateFormData(UpdateUserRolesSchema, { userId, roleNames })
    if (!validation.success) {
      return validation
    }

    const { data } = validation

    // Get current roles for audit log
    const { rows: userWithRoles } = await pool.query(
      `SELECT r.name FROM "Role" r JOIN "_RoleToUser" ru ON ru."A" = r.id WHERE ru."B" = $1`,
      [data.userId]
    )
    const oldRoles = userWithRoles.map((r: any) => r.name)

    // Get username for audit log
    const { rows: userRows } = await pool.query(
      `SELECT username FROM "User" WHERE id = $1`,
      [data.userId]
    )
    const currentUser = userRows[0] ?? null

    // Update roles using repository
    await userRepository.updateRoles(data.userId, data.roleNames)

    // Audit log
    await logDataChange(
      'update_user_roles',
      'user',
      data.userId,
      { roles: oldRoles },
      { roles: data.roleNames },
      { username: currentUser?.username }
    )

    logger.info('User roles updated', { userId, roleNames })
    revalidatePath('/admin')

    return { success: true as const }
  } catch (error) {
    logger.error('Failed to update user roles', { error, userId })
    return handleServerActionError(error)
  }
})

/**
 * Update user active status (Admin only)
 * Inactive users cannot sign in or access the site
 */
export const updateUserActiveStatus = withRole('admin', async (
  userId: string,
  isActive: boolean
) => {
  try {
    // Validate input
    const validation = validateFormData(UpdateUserActiveStatusSchema, { userId, isActive })
    if (!validation.success) {
      return validation
    }

    const { data } = validation

    // Get current user for audit log
    const { rows } = await pool.query(
      `SELECT username, "isActive" FROM "User" WHERE id = $1`,
      [data.userId]
    )
    const currentUser = rows[0] ?? null

    if (!currentUser) {
      return {
        success: false,
        error: 'User not found',
        code: 'NOT_FOUND'
      }
    }

    // Update user active status
    await pool.query(`UPDATE "User" SET "isActive" = $1 WHERE id = $2`, [data.isActive, data.userId])

    // Audit log
    await logDataChange(
      'update_user_active_status',
      'user',
      data.userId,
      { isActive: currentUser.isActive },
      { isActive: data.isActive },
      { username: currentUser.username }
    )

    logger.info('User active status updated', { userId, isActive })
    revalidatePath('/admin')

    return { success: true as const }
  } catch (error) {
    logger.error('Failed to update user active status', { error, userId })
    return handleServerActionError(error)
  }
})

/**
 * Update pupil information (Admin only)
 */
export const updatePupil = withRole('admin', async (
  admissionNumber: string,
  data: {
    firstName?: string
    lastName?: string
    gender?: string
    isActive?: boolean
  }
) => {
  try {
    // Validate input
    const validation = validateFormData(UpdatePupilSchema, { admissionNumber, data })
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    // Get current pupil for audit log
    const currentPupil = await pupilRepository.findByAdmissionNumber(validated.admissionNumber)

    // Update pupil using repository (handles encryption automatically)
    await pupilRepository.update(validated.admissionNumber, validated.data)

    // Audit log
    await logDataChange(
      'update_pupil',
      'pupil',
      validated.admissionNumber,
      currentPupil ? { firstName: currentPupil.firstName, lastName: currentPupil.lastName, gender: currentPupil.gender, isActive: currentPupil.isActive } : null,
      validated.data
    )

    logger.info('Pupil updated', { admissionNumber })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to update pupil', { error, admissionNumber })
    return handleServerActionError(error)
  }
})

/**
 * Get pupils with optional search query (Admin only)
 */
export const getPupils = withRole('admin', async (query: string = '') => {
  try {
    // Get pupils using repository (handles decryption automatically)
    const pupils = await pupilRepository.findAll(query)

    return { success: true, pupils }
  } catch (error) {
    logger.error('Failed to get pupils', { error, query })
    return handleServerActionError(error)
  }
})

/**
 * Process pupil upload from Excel/CSV (Admin only)
 */
export const processPupilUpload = withRole('admin', async (content: string) => {
  try {
    const workbook = XLSX.read(content, { type: 'base64' })
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json<any>(sheet)

    const pupils = rows.map((row: any) => ({
      admissionNumber: String(row.AdmissionNumber || row.admissionNumber || '').trim(),
      firstName: String(row.FirstName || row.firstName || row.Forename || '').trim(),
      lastName: String(row.LastName || row.lastName || row.Surname || '').trim(),
      gender: String(row.Gender || row.gender || 'M').trim().toUpperCase(),
      form: String(row.Form || row.form || '').trim() || null,
      isActive: true
    }))

    // Filter out invalid rows
    const validPupils = pupils.filter(p =>
      p.admissionNumber && p.firstName && p.lastName && (p.gender === 'M' || p.gender === 'F')
    )

    if (validPupils.length === 0) {
      return {
        success: false,
        error: 'No valid pupils found in upload',
        code: 'VALIDATION_ERROR'
      }
    }

    // Bulk create using repository (handles encryption automatically)
    const count = await pupilRepository.bulkCreate(validPupils)

    // Audit log
    await createAuditLog({
      action: 'upload_pupils',
      entityType: 'pupil',
      details: {
        totalInFile: validPupils.length,
        created: count,
        duplicatesSkipped: validPupils.length - count
      }
    })

    logger.info('Pupils uploaded', { count, total: validPupils.length })
    revalidatePath('/admin')

    return {
      success: true,
      message: `Successfully processed ${count} pupils (${validPupils.length - count} duplicates skipped)`
    }
  } catch (error) {
    logger.error('Failed to process pupil upload', { error })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Subject Management Actions
// ============================================================================

/**
 * Create a new subject (Admin only)
 */
export const createSubject = withRole('admin', async (formData: FormData) => {
  try {
    const data = {
      code: formData.get('code') as string,
      title: formData.get('title') as string,
      studiedComment: formData.get('introduction') as string
    }

    // Validate input
    const validation = validateFormData(CreateSubjectSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    // Create subject using repository
    const subject = await subjectRepository.create({
      code: validated.code,
      title: validated.title,
      studiedComment: validated.introduction
    })

    // Audit log
    await createAuditLog({
      action: 'create_subject',
      entityType: 'subject',
      entityId: subject.id,
      details: {
        after: { code: validated.code, title: validated.title }
      }
    })

    logger.info('Subject created', { code: validated.code })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to create subject', { error })
    return handleServerActionError(error)
  }
})

/**
 * Update a subject (Admin only)
 */
export const updateSubject = withRole('admin', async (
  subjectId: string,
  formData: FormData
) => {
  try {
    const data = {
      subjectId,
      code: formData.get('code') as string,
      title: formData.get('title') as string,
      studiedComment: formData.get('introduction') as string
    }

    // Validate input
    const validation = validateFormData(UpdateSubjectSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    // Get current subject for audit log
    const { rows: subjectRows } = await pool.query(`SELECT * FROM "Subject" WHERE id = $1`, [validated.subjectId])
    const currentSubject = subjectRows[0] ?? null

    // Update subject using repository
    await subjectRepository.update(validated.subjectId, {
      code: validated.code,
      title: validated.title,
      studiedComment: validated.introduction
    })

    // Audit log
    await logDataChange(
      'update_subject',
      'subject',
      validated.subjectId,
      currentSubject ? { code: currentSubject.code, title: currentSubject.title } : null,
      { code: validated.code, title: validated.title }
    )

    logger.info('Subject updated', { subjectId })
    revalidatePath('/admin')
    revalidatePath(`/hod/subject/${subjectId}`)

    return { success: true }
  } catch (error) {
    logger.error('Failed to update subject', { error, subjectId })
    return handleServerActionError(error)
  }
})

/**
 * Delete a subject (Admin only)
 */
export const deleteSubject = withRole('admin', async (subjectId: string) => {
  try {
    // Validate input
    const validation = validateFormData(DeleteSubjectSchema, { subjectId })
    if (!validation.success) {
      return validation
    }

    // Get current subject for audit log
    const { rows: subjectRows } = await pool.query(`SELECT * FROM "Subject" WHERE id = $1`, [subjectId])
    const currentSubject = subjectRows[0] ?? null

    // Delete subject using repository
    await subjectRepository.delete(subjectId)

    // Audit log
    await createAuditLog({
      action: 'delete_subject',
      entityType: 'subject',
      entityId: subjectId,
      details: currentSubject ? {
        before: { code: currentSubject.code, title: currentSubject.title }
      } : undefined
    })

    logger.info('Subject deleted', { subjectId })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete subject', { error, subjectId })
    return handleServerActionError(error)
  }
})

/**
 * Assign a user (HOD) to a subject (Admin only)
 */
export const assignUserToSubject = withRole('admin', async (
  subjectId: string,
  userId: string
) => {
  try {
    // Validate input
    const validation = validateFormData(AssignUserToSubjectSchema, { subjectId, userId })
    if (!validation.success) {
      return validation
    }

    // Get user and subject names for audit log
    const [{ rows: userRows }, { rows: subjectRows }] = await Promise.all([
      pool.query(`SELECT username FROM "User" WHERE id = $1`, [userId]),
      pool.query(`SELECT code, title FROM "Subject" WHERE id = $1`, [subjectId])
    ])
    const user = userRows[0] ?? null
    const subject = subjectRows[0] ?? null

    // Assign user using repository
    await subjectRepository.assignUser(subjectId, userId)

    // Audit log
    await createAuditLog({
      action: 'assign_user_to_subject',
      entityType: 'subject',
      entityId: subjectId,
      details: {
        userId,
        username: user?.username,
        subjectCode: subject?.code,
        subjectTitle: subject?.title
      }
    })

    logger.info('User assigned to subject', { subjectId, userId })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to assign user to subject', { error, subjectId, userId })
    return handleServerActionError(error)
  }
})

/**
 * Remove a user from a subject (Admin only)
 */
export const removeUserFromSubject = withRole('admin', async (
  subjectId: string,
  userId: string
) => {
  try {
    // Validate input
    const validation = validateFormData(AssignUserToSubjectSchema, { subjectId, userId })
    if (!validation.success) {
      return validation
    }

    // Get user and subject names for audit log
    const [{ rows: userRows }, { rows: subjectRows }] = await Promise.all([
      pool.query(`SELECT username FROM "User" WHERE id = $1`, [userId]),
      pool.query(`SELECT code, title FROM "Subject" WHERE id = $1`, [subjectId])
    ])
    const user = userRows[0] ?? null
    const subject = subjectRows[0] ?? null

    // Remove user using repository
    await subjectRepository.removeUser(subjectId, userId)

    // Audit log
    await createAuditLog({
      action: 'remove_user_from_subject',
      entityType: 'subject',
      entityId: subjectId,
      details: {
        userId,
        username: user?.username,
        subjectCode: subject?.code,
        subjectTitle: subject?.title
      }
    })

    logger.info('User removed from subject', { subjectId, userId })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to remove user from subject', { error, subjectId, userId })
    return handleServerActionError(error)
  }
})

/**
 * Assign teachers to a class (Admin only)
 */
export const assignTeachersToClass = withRole('admin', async (
  classId: string,
  teacherIds: string[]
) => {
  try {
    // Validate input
    const validation = validateFormData(AssignTeachersSchema, { classId, teacherIds })
    if (!validation.success) {
      return validation
    }

    // Get class info and current teachers for audit log
    const { rows: classRows } = await pool.query(`SELECT * FROM "Class" WHERE id = $1`, [classId])
    const { rows: currentTeachers } = await pool.query(
      `SELECT u.id, u.username FROM "User" u JOIN "_ClassToUser" cu ON cu."B" = u.id WHERE cu."A" = $1`,
      [classId]
    )
    const classInfo = classRows[0] ? { ...classRows[0], User: currentTeachers } : null
    const oldTeacherIds = classInfo?.User.map((u: any) => u.id) ?? []

    // Get new teacher usernames
    const { rows: newTeachers } = await pool.query(
      `SELECT id, username FROM "User" WHERE id = ANY($1::text[])`,
      [teacherIds]
    )

    // Assign teachers using repository
    await classRepository.assignTeachers(classId, teacherIds)

    // Audit log
    await logDataChange(
      'assign_teachers_to_class',
      'class',
      classId,
      { teacherIds: oldTeacherIds },
      { teacherIds },
      { className: classInfo?.name, teacherUsernames: newTeachers.map((t: any) => t.username) }
    )

    logger.info('Teachers assigned to class', { classId, teacherIds })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to assign teachers to class', { error, classId })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Class Management Actions
// ============================================================================

/**
 * Get all classes (Admin only)
 */
export const getClasses = withRole('admin', async () => {
  try {
    const classes = await classRepository.findAll()

    return { success: true as const, classes }
  } catch (error) {
    logger.error('Failed to get classes', { error })
    return handleServerActionError(error)
  }
})

/**
 * Create a class from a form (Admin only)
 * This creates a class and assigns all pupils from the specified form to it
 */
export const createClassFromForm = withRole('admin', async (
  className: string,
  formName: string,
  subjectId: string
) => {
  try {
    if (!className || !subjectId) {
      return {
        success: false,
        error: 'Class name and subject are required',
        code: 'VALIDATION_ERROR'
      }
    }

    // Extract year from form name or class name (e.g., "7A" -> "7")
    const year = (formName || className).match(/^\d+/)?.[0] || null

    // Create the class
    const newClass = await classRepository.create({
      name: className,
      year,
      subjectId
    })

    // If a form name is provided, assign all pupils from that form
    let pupilCount = 0
    if (formName) {
      const pupils = await pupilRepository.findByForm(formName)
      if (pupils.length > 0) {
        await classRepository.assignPupils(newClass.id, pupils.map(p => p.admissionNumber))
        pupilCount = pupils.length
      }
    }

    // Get subject info for audit log
    const { rows: subjectRows } = await pool.query(
      `SELECT code FROM "Subject" WHERE id = $1`,
      [subjectId]
    )
    const subject = subjectRows[0] ?? null

    // Audit log
    await createAuditLog({
      action: 'create_class',
      entityType: 'class',
      entityId: newClass.id,
      details: {
        after: { name: className, year, subjectCode: subject?.code },
        formName,
        pupilsAssigned: pupilCount
      }
    })

    logger.info('Class created', { className, formName, subjectId, hasFormPupils: !!formName })
    revalidatePath('/admin')

    return { success: true as const }
  } catch (error) {
    logger.error('Failed to create class', { error, className, formName, subjectId })
    return handleServerActionError(error)
  }
})

/**
 * Delete a class (Admin only)
 */
export const deleteClass = withRole('admin', async (classId: string) => {
  try {
    if (!classId) {
      return {
        success: false,
        error: 'Class ID is required',
        code: 'VALIDATION_ERROR'
      }
    }

    // Get class info for audit log (with subject code)
    const { rows: classRows } = await pool.query(`SELECT * FROM "Class" WHERE id = $1`, [classId])
    const classBase = classRows[0] ?? null
    let subjectCode: string | undefined
    if (classBase) {
      const { rows: subjectRows } = await pool.query(
        `SELECT code FROM "Subject" WHERE id = $1`,
        [classBase.subjectId]
      )
      subjectCode = subjectRows[0]?.code
    }
    const classInfo = classBase ? { ...classBase, Subject: subjectCode ? { code: subjectCode } : null } : null

    await classRepository.delete(classId)

    // Audit log
    await createAuditLog({
      action: 'delete_class',
      entityType: 'class',
      entityId: classId,
      details: classInfo ? {
        before: { name: classInfo.name, year: classInfo.year, subjectCode: classInfo.Subject?.code }
      } : undefined
    })

    logger.info('Class deleted', { classId })
    revalidatePath('/admin')

    return { success: true as const }
  } catch (error) {
    logger.error('Failed to delete class', { error, classId })
    return handleServerActionError(error)
  }
})

/**
 * Update a class (Admin only)
 */
export const updateClass = withRole('admin', async (
  classId: string,
  data: { name?: string; year?: string | null; subjectId?: string; teacherIds?: string[] }
) => {
  try {
    if (!classId) {
      return {
        success: false,
        error: 'Class ID is required',
        code: 'VALIDATION_ERROR'
      }
    }

    // Get current class info for audit log
    const { rows: classRows } = await pool.query(`SELECT * FROM "Class" WHERE id = $1`, [classId])
    const { rows: currentTeachers } = await pool.query(
      `SELECT u.id FROM "User" u JOIN "_ClassToUser" cu ON cu."B" = u.id WHERE cu."A" = $1`,
      [classId]
    )
    const currentClass = classRows[0]
      ? { ...classRows[0], User: currentTeachers }
      : null

    const { teacherIds, ...classData } = data

    // Update class details if any provided
    if (Object.keys(classData).length > 0) {
      await classRepository.update(classId, classData)
    }

    // Update teacher assignments if provided
    if (teacherIds !== undefined) {
      await classRepository.assignTeachers(classId, teacherIds)
    }

    // Audit log
    await logDataChange(
      'update_class',
      'class',
      classId,
      currentClass ? { name: currentClass.name, year: currentClass.year, teacherIds: currentClass.User.map((u: any) => u.id) } : null,
      { ...classData, ...(teacherIds !== undefined ? { teacherIds } : {}) }
    )

    logger.info('Class updated', { classId, data })
    revalidatePath('/admin')

    return { success: true as const }
  } catch (error) {
    logger.error('Failed to update class', { error, classId })
    return handleServerActionError(error)
  }
})

/**
 * Get all users (Admin only) - for assignment dropdowns
 */
export const getAllUsers = withRole('admin', async () => {
  try {
    const { rows } = await pool.query(
      `SELECT u.id, u.username, r.name as "roleName"
       FROM "User" u
       LEFT JOIN "_RoleToUser" ru ON ru."B" = u.id
       LEFT JOIN "Role" r ON r.id = ru."A"
       ORDER BY u.username ASC`
    )

    // Group by user in JS
    const userMap = new Map<string, { id: string; username: string; Role: { name: string }[] }>()
    for (const row of rows) {
      const existing = userMap.get(row.id)
      if (existing) {
        if (row.roleName) existing.Role.push({ name: row.roleName })
      } else {
        userMap.set(row.id, { id: row.id, username: row.username, Role: row.roleName ? [{ name: row.roleName }] : [] })
      }
    }
    const users = Array.from(userMap.values())

    return { success: true as const, users }
  } catch (error) {
    logger.error('Failed to get users', { error })
    return handleServerActionError(error)
  }
})

/**
 * Get pupils for a class (Admin only)
 */
export const getClassPupils = withRole('admin', async (classId: string) => {
  try {
    if (!classId) {
      return {
        success: false,
        error: 'Class ID is required',
        code: 'VALIDATION_ERROR'
      }
    }

    const pupils = await classRepository.getPupils(classId)
    return { success: true as const, pupils }
  } catch (error) {
    logger.error('Failed to get class pupils', { error, classId })
    return handleServerActionError(error)
  }
})

/**
 * Get all forms (distinct form values from pupils)
 */
export const getAllForms = withRole('admin', async () => {
  try {
    console.log('getAllForms: Starting query')
    const { rows } = await pool.query(
      `SELECT DISTINCT form FROM "Pupil" WHERE form IS NOT NULL AND "isActive" = true ORDER BY form ASC`
    )
    console.log('getAllForms: Query result', rows)

    const forms = rows.map((r: any) => r.form).filter(Boolean) as string[]
    console.log('getAllForms: Returning forms', forms)
    return { success: true as const, forms }
  } catch (error) {
    console.error('getAllForms: Error', error)
    logger.error('Failed to get forms', { error })
    return handleServerActionError(error)
  }
})

/**
 * Add pupils to a class (Admin only)
 */
export const addPupilsToClass = withRole('admin', async (classId: string, pupilAdmissionNumbers: string[]) => {
  try {
    if (!classId) {
      return {
        success: false,
        error: 'Class ID is required',
        code: 'VALIDATION_ERROR'
      }
    }

    // Get class info for audit log
    const { rows } = await pool.query(`SELECT * FROM "Class" WHERE id = $1`, [classId])
    const classInfo = rows[0] ?? null

    await classRepository.assignPupils(classId, pupilAdmissionNumbers)

    // Audit log
    await createAuditLog({
      action: 'add_pupils_to_class',
      entityType: 'class',
      entityId: classId,
      details: {
        className: classInfo?.name,
        pupilCount: pupilAdmissionNumbers.length,
        pupilAdmissionNumbers
      }
    })

    logger.info('Pupils added to class', { classId, count: pupilAdmissionNumbers.length })
    revalidatePath('/admin')

    return { success: true as const }
  } catch (error) {
    logger.error('Failed to add pupils to class', { error, classId })
    return handleServerActionError(error)
  }
})

/**
 * Remove pupils from a class (Admin only)
 */
export const removePupilsFromClass = withRole('admin', async (classId: string, pupilAdmissionNumbers: string[]) => {
  try {
    if (!classId) {
      return {
        success: false,
        error: 'Class ID is required',
        code: 'VALIDATION_ERROR'
      }
    }

    // Get class info for audit log
    const { rows } = await pool.query(`SELECT * FROM "Class" WHERE id = $1`, [classId])
    const classInfo = rows[0] ?? null

    await classRepository.removePupils(classId, pupilAdmissionNumbers)

    // Audit log
    await createAuditLog({
      action: 'remove_pupils_from_class',
      entityType: 'class',
      entityId: classId,
      details: {
        className: classInfo?.name,
        pupilCount: pupilAdmissionNumbers.length,
        pupilAdmissionNumbers
      }
    })

    logger.info('Pupils removed from class', { classId, count: pupilAdmissionNumbers.length })
    revalidatePath('/admin')

    return { success: true as const }
  } catch (error) {
    logger.error('Failed to remove pupils from class', { error, classId })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Deadline Management Actions
// ============================================================================

/**
 * Get all deadlines (Admin only)
 */
export const getDeadlines = withRole('admin', async () => {
  try {
    const { rows: deadlines } = await pool.query(
      `SELECT * FROM "Deadline" ORDER BY date ASC`
    )
    return { success: true as const, deadlines }
  } catch (error) {
    logger.error('Failed to get deadlines', { error })
    return handleServerActionError(error)
  }
})

/**
 * Create a new deadline (Admin only)
 */
export const createDeadline = withRole('admin', async (formData: FormData) => {
  try {
    const data = {
      title: formData.get('title') as string,
      date: formData.get('date') as string,
      description: formData.get('description') as string || undefined
    }

    const validation = validateFormData(CreateDeadlineSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data
    const createdId = createId()
    await pool.query(
      `INSERT INTO "Deadline" (id, title, date, description) VALUES ($1, $2, $3, $4)`,
      [createdId, validated.title, new Date(validated.date), validated.description ?? null]
    )
    const deadline = { id: createdId }

    // Audit log
    await createAuditLog({
      action: 'create_deadline',
      entityType: 'deadline',
      entityId: deadline.id,
      details: {
        after: { title: validated.title, date: validated.date }
      }
    })

    logger.info('Deadline created', { title: validated.title })
    revalidatePath('/admin')
    revalidatePath('/')

    return { success: true }
  } catch (error) {
    logger.error('Failed to create deadline', { error })
    return handleServerActionError(error)
  }
})

/**
 * Update a deadline (Admin only)
 */
export const updateDeadline = withRole('admin', async (
  deadlineId: string,
  formData: FormData
) => {
  try {
    const data = {
      deadlineId,
      title: formData.get('title') as string || undefined,
      date: formData.get('date') as string || undefined,
      description: formData.get('description') as string || undefined,
      isActive: formData.has('isActive') ? formData.get('isActive') === 'true' : undefined
    }

    const validation = validateFormData(UpdateDeadlineSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    // Get current deadline for audit log
    const { rows: deadlineRows } = await pool.query(`SELECT * FROM "Deadline" WHERE id = $1`, [validated.deadlineId])
    const currentDeadline = deadlineRows[0] ?? null

    const sets: string[] = []
    const params: any[] = []
    let idx = 1

    if (validated.title !== undefined) { sets.push(`title = $${idx++}`); params.push(validated.title) }
    if (validated.date !== undefined) { sets.push(`date = $${idx++}`); params.push(new Date(validated.date)) }
    if (validated.description !== undefined) { sets.push(`description = $${idx++}`); params.push(validated.description) }
    if (validated.isActive !== undefined) { sets.push(`"isActive" = $${idx++}`); params.push(validated.isActive) }

    if (sets.length === 0) {
      // nothing to update — return success without hitting DB
      return { success: true }
    }

    params.push(validated.deadlineId)
    await pool.query(
      `UPDATE "Deadline" SET ${sets.join(', ')} WHERE id = $${idx}`,
      params
    )

    // Audit log
    await logDataChange(
      'update_deadline',
      'deadline',
      validated.deadlineId,
      currentDeadline ? { title: currentDeadline.title, date: currentDeadline.date instanceof Date ? currentDeadline.date.toISOString() : currentDeadline.date, isActive: currentDeadline.isActive } : null,
      { title: validated.title, date: validated.date, isActive: validated.isActive }
    )

    logger.info('Deadline updated', { deadlineId })
    revalidatePath('/admin')
    revalidatePath('/')

    return { success: true }
  } catch (error) {
    logger.error('Failed to update deadline', { error, deadlineId })
    return handleServerActionError(error)
  }
})

/**
 * Delete a deadline (Admin only)
 */
export const deleteDeadline = withRole('admin', async (deadlineId: string) => {
  try {
    const validation = validateFormData(DeleteDeadlineSchema, { deadlineId })
    if (!validation.success) {
      return validation
    }

    // Get current deadline for audit log
    const { rows: deadlineRows } = await pool.query(`SELECT * FROM "Deadline" WHERE id = $1`, [deadlineId])
    const currentDeadline = deadlineRows[0] ?? null

    await pool.query(`DELETE FROM "Deadline" WHERE id = $1`, [deadlineId])

    // Audit log
    await createAuditLog({
      action: 'delete_deadline',
      entityType: 'deadline',
      entityId: deadlineId,
      details: {
        before: currentDeadline ? { title: currentDeadline.title, date: currentDeadline.date instanceof Date ? currentDeadline.date.toISOString() : currentDeadline.date } : undefined
      }
    })

    logger.info('Deadline deleted', { deadlineId })
    revalidatePath('/admin')
    revalidatePath('/')

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete deadline', { error, deadlineId })
    return handleServerActionError(error)
  }
})
