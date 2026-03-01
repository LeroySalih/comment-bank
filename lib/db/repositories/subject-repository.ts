import { pool } from '@/lib/db'
import { NotFoundError } from '@/lib/errors'
import { DbSubject, DbClass, DbCommentGroup, DbCommentOption } from '@/lib/types/db'

interface SubjectHod {
  id: string
  username: string
}

interface SubjectWithBasicRelations extends DbSubject {
  Class: DbClass[]
  User: SubjectHod[]
  _count: { CommentGroup: number }
}

interface CommentGroupWithOptions extends DbCommentGroup {
  CommentOption: DbCommentOption[]
}

interface SubjectWithFullDetails extends DbSubject {
  Class: DbClass[]
  User: SubjectHod[]
  CommentGroup: CommentGroupWithOptions[]
}

interface SubjectWithRelations extends DbSubject {
  Class: DbClass[]
  User: SubjectHod[]
}

/**
 * Repository for Subject data access
 */
export class SubjectRepository {
  /**
   * Find all subjects with basic relations
   */
  async findAll(): Promise<SubjectWithBasicRelations[]> {
    const { rows: subjects } = await pool.query<DbSubject>(
      `SELECT * FROM "Subject" ORDER BY code ASC`
    )

    const results: SubjectWithBasicRelations[] = []

    for (const subject of subjects) {
      const [classesResult, hodsResult, countResult] = await Promise.all([
        pool.query<DbClass>(
          `SELECT * FROM "Class" WHERE "subjectId" = $1 ORDER BY name ASC`,
          [subject.id]
        ),
        pool.query<SubjectHod>(
          `SELECT u.id, u.username FROM "User" u JOIN "_SubjectToUser" su ON su."B" = u.id WHERE su."A" = $1`,
          [subject.id]
        ),
        pool.query<{ count: string }>(
          `SELECT COUNT(*) as count FROM "CommentGroup" WHERE "subjectId" = $1`,
          [subject.id]
        )
      ])

      results.push({
        ...subject,
        Class: classesResult.rows,
        User: hodsResult.rows,
        _count: { CommentGroup: parseInt(countResult.rows[0].count, 10) }
      })
    }

    return results
  }

  /**
   * Find subject by ID with full details
   */
  async findByIdWithDetails(id: string): Promise<SubjectWithFullDetails> {
    const { rows } = await pool.query<DbSubject>(
      `SELECT * FROM "Subject" WHERE id = $1`,
      [id]
    )

    if (rows.length === 0) {
      throw new NotFoundError(`Subject with ID ${id} not found`)
    }

    const subject = rows[0]

    const [classesResult, hodsResult, groupsResult] = await Promise.all([
      pool.query<DbClass>(
        `SELECT * FROM "Class" WHERE "subjectId" = $1 ORDER BY name ASC`,
        [id]
      ),
      pool.query<SubjectHod>(
        `SELECT u.id, u.username FROM "User" u JOIN "_SubjectToUser" su ON su."B" = u.id WHERE su."A" = $1`,
        [id]
      ),
      pool.query<DbCommentGroup>(
        `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 ORDER BY "displayOrder" ASC`,
        [id]
      )
    ])

    const commentGroups: CommentGroupWithOptions[] = []

    for (const group of groupsResult.rows) {
      const optionsResult = await pool.query<DbCommentOption>(
        `SELECT * FROM "CommentOption" WHERE "groupId" = $1 ORDER BY "displayOrder" ASC`,
        [group.id]
      )
      commentGroups.push({ ...group, CommentOption: optionsResult.rows })
    }

    return {
      ...subject,
      Class: classesResult.rows,
      User: hodsResult.rows,
      CommentGroup: commentGroups
    }
  }

  /**
   * Find subject by code
   */
  async findByCode(code: string): Promise<DbSubject | null> {
    const { rows } = await pool.query<DbSubject>(
      `SELECT * FROM "Subject" WHERE code = $1`,
      [code]
    )
    return rows[0] ?? null
  }

  /**
   * Create a new subject
   */
  async create(data: {
    code: string
    title?: string
    studiedComment?: string
  }): Promise<SubjectWithRelations> {
    const { rows } = await pool.query<DbSubject>(
      `INSERT INTO "Subject" (id, code, title, "studiedComment") VALUES ($1, $2, $3, $4) RETURNING *`,
      [data.code, data.code, data.title ?? null, data.studiedComment ?? null]
    )
    return { ...rows[0], Class: [], User: [] }
  }

  /**
   * Update a subject
   */
  async update(
    id: string,
    data: {
      code?: string
      title?: string
      studiedComment?: string
    }
  ): Promise<SubjectWithRelations> {
    const fields: string[] = []
    const values: unknown[] = []
    let paramIndex = 1

    if (data.code !== undefined) {
      fields.push(`code = $${paramIndex++}`)
      values.push(data.code)
    }
    if (data.title !== undefined) {
      fields.push(`title = $${paramIndex++}`)
      values.push(data.title)
    }
    if (data.studiedComment !== undefined) {
      fields.push(`"studiedComment" = $${paramIndex++}`)
      values.push(data.studiedComment)
    }

    values.push(id)

    const { rows } = await pool.query<DbSubject>(
      `UPDATE "Subject" SET ${fields.join(', ')} WHERE id = $${paramIndex} RETURNING *`,
      values
    )

    return { ...rows[0], Class: [], User: [] }
  }

  /**
   * Delete a subject
   */
  async delete(id: string): Promise<void> {
    await pool.query(`DELETE FROM "Subject" WHERE id = $1`, [id])
  }

  /**
   * Assign a user (HOD) to a subject
   */
  async assignUser(subjectId: string, userId: string): Promise<{ User: SubjectHod[] }> {
    await pool.query(
      `INSERT INTO "_SubjectToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [subjectId, userId]
    )

    const { rows } = await pool.query<SubjectHod>(
      `SELECT u.id, u.username FROM "User" u JOIN "_SubjectToUser" su ON su."B" = u.id WHERE su."A" = $1`,
      [subjectId]
    )

    return { User: rows }
  }

  /**
   * Remove a user from a subject
   */
  async removeUser(subjectId: string, userId: string): Promise<{ User: SubjectHod[] }> {
    await pool.query(
      `DELETE FROM "_SubjectToUser" WHERE "A" = $1 AND "B" = $2`,
      [subjectId, userId]
    )

    const { rows } = await pool.query<SubjectHod>(
      `SELECT u.id, u.username FROM "User" u JOIN "_SubjectToUser" su ON su."B" = u.id WHERE su."A" = $1`,
      [subjectId]
    )

    return { User: rows }
  }
}

export const subjectRepository = new SubjectRepository()
