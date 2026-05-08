import { pool } from '@/lib/db'
import { decrypt } from '@/lib/encryption'
import { NotFoundError, ForbiddenError } from '@/lib/errors'
import { createId } from '@paralleldrive/cuid2'
import type {
  DbClass,
  DbSubject,
  DbCommentGroup,
  DbCommentOption,
  DbAssignment,
  DbPupil,
  DbPupilCode
} from '@/lib/types/db'

// ---------------------------------------------------------------------------
// Local helper interfaces
// ---------------------------------------------------------------------------

interface ClassTeacher {
  id: string
  username: string
}

interface ClassWithRelations extends DbClass {
  Subject: DbSubject
  User: ClassTeacher[]
}

interface CommentGroupWithOptions extends DbCommentGroup {
  CommentOption: DbCommentOption[]
}

interface SubjectWithCommentGroups extends DbSubject {
  CommentGroup: CommentGroupWithOptions[]
}

interface PupilCodeWithCommentGroup extends DbPupilCode {
  CommentGroup: DbCommentGroup
}

interface AssignmentWithPupilAndCodes extends DbAssignment {
  Pupil: DbPupil
  PupilCode: PupilCodeWithCommentGroup[]
}

interface ClassWithAssignments extends DbClass {
  Subject: SubjectWithCommentGroups
  User: ClassTeacher[]
  Assignment: AssignmentWithPupilAndCodes[]
}

interface ClassWithCount extends DbClass {
  Subject: DbSubject
  User: ClassTeacher[]
  _count: { Assignment: number }
}

interface ClassWithSubjectAndTeachers extends DbClass {
  Subject: Pick<DbSubject, 'code' | 'title'>
  User: ClassTeacher[]
}

// Raw row types for SQL results
interface AssignmentRawRow extends DbAssignment {
  pupil_admissionNumber: string
  pupil_firstName: string
  pupil_lastName: string
  pupil_gender: string
  pupil_isActive: boolean
  pupil_form: string | null
}

interface PupilCodeRawRow extends DbPupilCode {
  cg_id: string
  cg_name: string
  cg_displayOrder: number
  cg_subjectId: string
  cg_title: string
  cg_isLinked: boolean
  cg_linkedField: string | null
}

interface PupilWithAssignmentId extends DbPupil {
  assignmentId: string
}

/**
 * Repository for Class data access
 */
export class ClassRepository {
  /**
   * Find class by ID with basic info (subject + teachers)
   */
  async findById(id: string): Promise<ClassWithRelations> {
    const { rows: classRows } = await pool.query<DbClass>(
      `SELECT * FROM "Class" WHERE id = $1`,
      [id]
    )

    if (classRows.length === 0) {
      throw new NotFoundError(`Class with ID ${id} not found`)
    }

    const cls = classRows[0]

    const [subjectResult, teachersResult] = await Promise.all([
      pool.query<DbSubject>(`SELECT * FROM "Subject" WHERE id = $1`, [cls.subjectId]),
      pool.query<ClassTeacher>(
        `SELECT u.id, u.username FROM "User" u
         JOIN "_ClassToUser" cu ON cu."B" = u.id
         WHERE cu."A" = $1`,
        [id]
      )
    ])

    return {
      ...cls,
      Subject: subjectResult.rows[0],
      User: teachersResult.rows
    }
  }

  /**
   * Find class by ID with full assignment details.
   * Automatically decrypts pupil PII.
   */
  async findByIdWithAssignments(id: string): Promise<ClassWithAssignments> {
    // Step 1: get class + subject + teachers
    const cls = await this.findById(id)

    // Step 2: fetch comment groups for the subject, then options per group
    const { rows: groupRows } = await pool.query<DbCommentGroup>(
      `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 ORDER BY "displayOrder" ASC`,
      [cls.subjectId]
    )

    const commentGroups: CommentGroupWithOptions[] = []
    for (const group of groupRows) {
      const { rows: optionRows } = await pool.query<DbCommentOption>(
        `SELECT * FROM "CommentOption" WHERE "groupId" = $1 ORDER BY "displayOrder" ASC`,
        [group.id]
      )
      commentGroups.push({ ...group, CommentOption: optionRows })
    }

    const subjectWithGroups: SubjectWithCommentGroups = {
      ...cls.Subject,
      CommentGroup: commentGroups
    }

    // Step 3: fetch assignments with pupil info (active pupils only, ordered by lastName)
    const { rows: assignmentRows } = await pool.query<AssignmentRawRow>(
      `SELECT a.*,
         p."admissionNumber" as pupil_admissionNumber,
         p."firstName"       as pupil_firstName,
         p."lastName"        as pupil_lastName,
         p."gender"          as pupil_gender,
         p."isActive"        as pupil_isActive,
         p."form"            as pupil_form
       FROM "Assignment" a
       JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
       WHERE a."classId" = $1 AND p."isActive" = true
       ORDER BY p."lastName" ASC`,
      [id]
    )

    // Step 4: fetch all PupilCode rows for those assignments (with CommentGroup)
    let pupilCodesByAssignment = new Map<string, PupilCodeWithCommentGroup[]>()

    if (assignmentRows.length > 0) {
      const assignmentIds = assignmentRows.map(r => r.id)

      const { rows: pupilCodeRows } = await pool.query<PupilCodeRawRow>(
        `SELECT pc.*,
           cg.id              as cg_id,
           cg.name            as cg_name,
           cg."displayOrder"  as cg_displayOrder,
           cg."subjectId"     as cg_subjectId,
           cg.title           as cg_title,
           cg."isLinked"      as cg_isLinked,
           cg."linkedField"   as cg_linkedField
         FROM "PupilCode" pc
         JOIN "CommentGroup" cg ON cg.id = pc."groupId"
         WHERE pc."assignmentId" = ANY($1::text[])`,
        [assignmentIds]
      )

      // Step 5: group PupilCodes by assignmentId
      for (const row of pupilCodeRows) {
        const commentGroup: DbCommentGroup = {
          id: row.cg_id,
          name: row.cg_name,
          displayOrder: row.cg_displayOrder,
          subjectId: row.cg_subjectId,
          title: row.cg_title,
          isLinked: row.cg_isLinked,
          linkedField: row.cg_linkedField
        }

        const pupilCode: PupilCodeWithCommentGroup = {
          id: row.id,
          assignmentId: row.assignmentId,
          groupId: row.groupId,
          code: row.code,
          CommentGroup: commentGroup
        }

        const existing = pupilCodesByAssignment.get(row.assignmentId) ?? []
        existing.push(pupilCode)
        pupilCodesByAssignment.set(row.assignmentId, existing)
      }
    }

    // Step 5 (cont): reshape assignments — decrypt PII and attach PupilCodes
    const assignments: AssignmentWithPupilAndCodes[] = assignmentRows.map(row => {
      const pupil: DbPupil = {
        admissionNumber: row.pupil_admissionNumber,
        firstName: decrypt(row.pupil_firstName),
        lastName: decrypt(row.pupil_lastName),
        gender: row.pupil_gender,
        isActive: row.pupil_isActive,
        form: row.pupil_form
      }

      const assignmentBase: DbAssignment = {
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
        checkedById: row.checkedById
      }

      return {
        ...assignmentBase,
        Pupil: pupil,
        PupilCode: pupilCodesByAssignment.get(row.id) ?? []
      }
    })

    return {
      ...cls,
      Subject: subjectWithGroups,
      Assignment: assignments
    }
  }

  /**
   * Find all classes for a specific teacher
   */
  async findByTeacherId(teacherId: string): Promise<ClassWithCount[]> {
    const { rows: classRows } = await pool.query<DbClass>(
      `SELECT c.* FROM "Class" c
       JOIN "_ClassToUser" cu ON cu."A" = c.id
       WHERE cu."B" = $1
       ORDER BY c.name ASC`,
      [teacherId]
    )

    const results: ClassWithCount[] = []

    for (const cls of classRows) {
      const [subjectResult, teachersResult, countResult] = await Promise.all([
        pool.query<DbSubject>(`SELECT * FROM "Subject" WHERE id = $1`, [cls.subjectId]),
        pool.query<ClassTeacher>(
          `SELECT u.id, u.username FROM "User" u
           JOIN "_ClassToUser" cu ON cu."B" = u.id WHERE cu."A" = $1`,
          [cls.id]
        ),
        pool.query<{ count: string }>(
          `SELECT COUNT(*) as count FROM "Assignment" WHERE "classId" = $1`,
          [cls.id]
        )
      ])

      results.push({
        ...cls,
        Subject: subjectResult.rows[0],
        User: teachersResult.rows,
        _count: { Assignment: parseInt(countResult.rows[0].count, 10) }
      })
    }

    return results
  }

  /**
   * Assign teachers to a class (replaces all existing teachers in a transaction)
   */
  async assignTeachers(classId: string, teacherIds: string[]): Promise<{ User: ClassTeacher[] }> {
    const client = await pool.connect()
    try {
      await client.query('BEGIN')

      await client.query(`DELETE FROM "_ClassToUser" WHERE "A" = $1`, [classId])

      for (const teacherId of teacherIds) {
        await client.query(
          `INSERT INTO "_ClassToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [classId, teacherId]
        )
      }

      await client.query('COMMIT')
    } catch (err) {
      await client.query('ROLLBACK')
      throw err
    } finally {
      client.release()
    }

    const { rows } = await pool.query<ClassTeacher>(
      `SELECT u.id, u.username FROM "User" u
       JOIN "_ClassToUser" cu ON cu."B" = u.id WHERE cu."A" = $1`,
      [classId]
    )

    return { User: rows }
  }

  /**
   * Check if a user has access to a class.
   * Admins always have access. HODs have access to classes in their subjects.
   * Teachers must be assigned to the class.
   */
  async authorizeAccess(classId: string, userId: string, userRoles: string[]): Promise<boolean> {
    if (userRoles.includes('admin')) {
      return true
    }

    const cls = await this.findById(classId)

    if (userRoles.includes('hod')) {
      const { rows } = await pool.query(
        `SELECT 1 FROM "_SubjectToUser" WHERE "A" = $1 AND "B" = $2 LIMIT 1`,
        [cls.subjectId, userId]
      )
      if (rows.length > 0) {
        return true
      }
    }

    return cls.User.some(teacher => teacher.id === userId)
  }

  /**
   * Require access to a class (throws ForbiddenError if unauthorized)
   */
  async requireAccess(classId: string, userId: string, userRoles: string[]): Promise<void> {
    const hasAccess = await this.authorizeAccess(classId, userId, userRoles)
    if (!hasAccess) {
      throw new ForbiddenError('You do not have access to this class')
    }
  }

  /**
   * Find all classes with subject, teachers, and assignment count
   */
  async findAll(): Promise<ClassWithCount[]> {
    const { rows: classRows } = await pool.query<DbClass>(
      `SELECT * FROM "Class" ORDER BY name ASC`
    )

    const results: ClassWithCount[] = []

    for (const cls of classRows) {
      const [subjectResult, teachersResult, countResult] = await Promise.all([
        pool.query<DbSubject>(`SELECT * FROM "Subject" WHERE id = $1`, [cls.subjectId]),
        pool.query<ClassTeacher>(
          `SELECT u.id, u.username FROM "User" u
           JOIN "_ClassToUser" cu ON cu."B" = u.id WHERE cu."A" = $1`,
          [cls.id]
        ),
        pool.query<{ count: string }>(
          `SELECT COUNT(*) as count FROM "Assignment" WHERE "classId" = $1`,
          [cls.id]
        )
      ])

      results.push({
        ...cls,
        Subject: subjectResult.rows[0],
        User: teachersResult.rows,
        _count: { Assignment: parseInt(countResult.rows[0].count, 10) }
      })
    }

    return results
  }

  /**
   * Create a new class
   */
  async create(data: { name: string; year: string | null; subjectId: string }): Promise<DbClass> {
    const { rows } = await pool.query<DbClass>(
      `INSERT INTO "Class" (id, name, year, "subjectId") VALUES ($1, $2, $3, $4) RETURNING *`,
      [createId(), data.name, data.year, data.subjectId]
    )
    return rows[0]
  }

  /**
   * Update a class (dynamic SET), returns with subject and teachers attached
   */
  async update(
    classId: string,
    data: { name?: string; year?: string | null; subjectId?: string }
  ): Promise<ClassWithSubjectAndTeachers> {
    const fields: string[] = []
    const values: unknown[] = []
    let paramIndex = 1

    if (data.name !== undefined) {
      fields.push(`name = $${paramIndex++}`)
      values.push(data.name)
    }
    if (data.year !== undefined) {
      fields.push(`year = $${paramIndex++}`)
      values.push(data.year)
    }
    if (data.subjectId !== undefined) {
      fields.push(`"subjectId" = $${paramIndex++}`)
      values.push(data.subjectId)
    }

    if (fields.length === 0) {
      throw new Error('update called with no fields to update')
    }

    values.push(classId)

    const { rows } = await pool.query<DbClass>(
      `UPDATE "Class" SET ${fields.join(', ')} WHERE id = $${paramIndex} RETURNING *`,
      values
    )

    const cls = rows[0]

    const [subjectResult, teachersResult] = await Promise.all([
      pool.query<DbSubject>(`SELECT * FROM "Subject" WHERE id = $1`, [cls.subjectId]),
      pool.query<ClassTeacher>(
        `SELECT u.id, u.username FROM "User" u
         JOIN "_ClassToUser" cu ON cu."B" = u.id WHERE cu."A" = $1`,
        [cls.id]
      )
    ])

    return {
      ...cls,
      Subject: { code: subjectResult.rows[0].code, title: subjectResult.rows[0].title },
      User: teachersResult.rows
    }
  }

  /**
   * Delete a class
   */
  async delete(classId: string): Promise<void> {
    await pool.query(`DELETE FROM "Class" WHERE id = $1`, [classId])
  }

  /**
   * Assign pupils to a class by bulk-inserting assignments (skips duplicates)
   */
  async assignPupils(
    classId: string,
    pupilAdmissionNumbers: string[]
  ): Promise<{ count: number }> {
    if (pupilAdmissionNumbers.length === 0) {
      return { count: 0 }
    }

    const params: unknown[] = []
    const valuePlaceholders = pupilAdmissionNumbers.map(admissionNumber => {
      const base = params.length
      params.push(`${classId}-${admissionNumber}`, classId, admissionNumber)
      return `($${base + 1}, $${base + 2}, $${base + 3})`
    })

    const { rowCount } = await pool.query(
      `INSERT INTO "Assignment" (id, "classId", "pupilId")
       VALUES ${valuePlaceholders.join(', ')}
       ON CONFLICT DO NOTHING`,
      params
    )

    return { count: rowCount ?? 0 }
  }

  /**
   * Get pupils assigned to a class.
   * Automatically decrypts PII.
   */
  async getPupils(classId: string): Promise<PupilWithAssignmentId[]> {
    const { rows } = await pool.query<DbPupil & { assignmentId: string; finalComment: string | null }>(
      `SELECT p.*, a.id as "assignmentId", a."finalComment"
       FROM "Assignment" a
       JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
       WHERE a."classId" = $1
       ORDER BY p."lastName" ASC`,
      [classId]
    )

    return rows.map(row => ({
      ...row,
      firstName: decrypt(row.firstName),
      lastName: decrypt(row.lastName)
    }))
  }

  /**
   * Remove a single pupil from a class
   */
  async removePupil(
    classId: string,
    pupilAdmissionNumber: string
  ): Promise<{ count: number }> {
    const { rowCount } = await pool.query(
      `DELETE FROM "Assignment" WHERE "classId" = $1 AND "pupilId" = $2`,
      [classId, pupilAdmissionNumber]
    )
    return { count: rowCount ?? 0 }
  }

  /**
   * Remove multiple pupils from a class
   */
  async removePupils(
    classId: string,
    pupilAdmissionNumbers: string[]
  ): Promise<{ count: number }> {
    const { rowCount } = await pool.query(
      `DELETE FROM "Assignment" WHERE "classId" = $1 AND "pupilId" = ANY($2::text[])`,
      [classId, pupilAdmissionNumbers]
    )
    return { count: rowCount ?? 0 }
  }
}

export const classRepository = new ClassRepository()
