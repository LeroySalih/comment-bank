import { pool } from '@/lib/db'
import { encrypt, decrypt } from '@/lib/encryption'
import { NotFoundError } from '@/lib/errors'
import type { DbPupil } from '@/lib/types/db'

/**
 * Repository for Pupil data access
 * Handles encryption/decryption automatically
 */
export class PupilRepository {
  /**
   * Find all pupils with optional search query
   * Automatically decrypts PII fields
   */
  async findAll(query?: string): Promise<DbPupil[]> {
    let result

    if (query) {
      const param = `%${query}%`
      result = await pool.query<DbPupil>(
        `SELECT * FROM "Pupil"
         WHERE "firstName" ILIKE $1 OR "lastName" ILIKE $1 OR "admissionNumber" ILIKE $1
         ORDER BY "lastName" ASC, "firstName" ASC`,
        [param]
      )
    } else {
      result = await pool.query<DbPupil>(
        `SELECT * FROM "Pupil" ORDER BY "lastName" ASC, "firstName" ASC`
      )
    }

    return result.rows.map(pupil => ({
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }))
  }

  /**
   * Find pupil by admission number
   * Automatically decrypts PII fields
   */
  async findByAdmissionNumber(admissionNumber: string): Promise<DbPupil | null> {
    const { rows } = await pool.query<DbPupil>(
      `SELECT * FROM "Pupil" WHERE "admissionNumber" = $1`,
      [admissionNumber]
    )

    if (rows.length === 0) return null

    const pupil = rows[0]
    return {
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }
  }

  /**
   * Create a new pupil
   * Automatically encrypts PII fields
   */
  async create(data: {
    admissionNumber: string
    firstName: string
    lastName: string
    gender: string
    form?: string | null
    isActive?: boolean
  }): Promise<DbPupil> {
    const { rows } = await pool.query<DbPupil>(
      `INSERT INTO "Pupil" ("admissionNumber", "firstName", "lastName", "gender", "form", "isActive")
       VALUES ($1, $2, $3, $4, $5, $6) RETURNING *`,
      [
        data.admissionNumber,
        encrypt(data.firstName),
        encrypt(data.lastName),
        data.gender,
        data.form ?? null,
        data.isActive ?? true
      ]
    )

    const pupil = rows[0]
    return {
      ...pupil,
      firstName: data.firstName, // Return unencrypted
      lastName: data.lastName
    }
  }

  /**
   * Update a pupil
   * Automatically encrypts PII fields if provided
   */
  async update(
    admissionNumber: string,
    data: {
      firstName?: string
      lastName?: string
      gender?: string
      isActive?: boolean
    }
  ): Promise<DbPupil> {
    const setClauses: string[] = []
    const params: unknown[] = []

    if (data.firstName !== undefined) {
      params.push(encrypt(data.firstName))
      setClauses.push(`"firstName" = $${params.length}`)
    }
    if (data.lastName !== undefined) {
      params.push(encrypt(data.lastName))
      setClauses.push(`"lastName" = $${params.length}`)
    }
    if (data.gender !== undefined) {
      params.push(data.gender)
      setClauses.push(`"gender" = $${params.length}`)
    }
    if (data.isActive !== undefined) {
      params.push(data.isActive)
      setClauses.push(`"isActive" = $${params.length}`)
    }

    if (setClauses.length === 0) {
      throw new Error('update called with no fields to update')
    }

    params.push(admissionNumber)
    const whereParam = `$${params.length}`

    const { rows } = await pool.query<DbPupil>(
      `UPDATE "Pupil" SET ${setClauses.join(', ')} WHERE "admissionNumber" = ${whereParam} RETURNING *`,
      params
    )

    if (rows.length === 0) {
      throw new NotFoundError(`Pupil with admissionNumber ${admissionNumber} not found`)
    }

    const pupil = rows[0]
    return {
      ...pupil,
      firstName: data.firstName !== undefined ? data.firstName : decrypt(pupil.firstName),
      lastName: data.lastName !== undefined ? data.lastName : decrypt(pupil.lastName)
    }
  }

  /**
   * Bulk create pupils
   * Automatically encrypts PII fields
   */
  async bulkCreate(
    pupils: Array<{
      admissionNumber: string
      firstName: string
      lastName: string
      gender: string
      form?: string | null
      isActive?: boolean
    }>
  ): Promise<number> {
    if (pupils.length === 0) return 0

    const params: unknown[] = []
    const valuePlaceholders = pupils.map(pupil => {
      const base = params.length
      params.push(
        pupil.admissionNumber,
        encrypt(pupil.firstName),
        encrypt(pupil.lastName),
        pupil.gender,
        pupil.form ?? null,
        pupil.isActive ?? true
      )
      return `($${base + 1}, $${base + 2}, $${base + 3}, $${base + 4}, $${base + 5}, $${base + 6})`
    })

    const { rowCount } = await pool.query(
      `INSERT INTO "Pupil" ("admissionNumber", "firstName", "lastName", "gender", "form", "isActive")
       VALUES ${valuePlaceholders.join(', ')}
       ON CONFLICT ("admissionNumber") DO NOTHING`,
      params
    )

    return rowCount ?? 0
  }

  /**
   * Find pupils by form name
   * Automatically decrypts PII fields
   */
  async findByForm(formName: string): Promise<DbPupil[]> {
    const { rows } = await pool.query<DbPupil>(
      `SELECT * FROM "Pupil" WHERE "form" = $1 AND "isActive" = true
       ORDER BY "lastName" ASC, "firstName" ASC`,
      [formName]
    )

    return rows.map(pupil => ({
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }))
  }
}


export const pupilRepository = new PupilRepository()
