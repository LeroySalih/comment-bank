import { pool } from '@/lib/db'
import { NotFoundError } from '@/lib/errors'
import { createId } from '@paralleldrive/cuid2'
import type { DbUser, DbRole, DbUserWithRoles, DbSubject, DbClass } from '@/lib/types/db'

type DbUserWithRelations = DbUserWithRoles & {
  Subject: DbSubject[]
  Class: DbClass[]
}

/**
 * Repository for User data access
 */
export class UserRepository {
  /**
   * Find all users with their roles
   */
  async findAll(): Promise<DbUserWithRoles[]> {
    const usersResult = await pool.query<DbUser>(
      `SELECT * FROM "User" ORDER BY "username" ASC`
    )
    const users = usersResult.rows

    const usersWithRoles = await Promise.all(
      users.map(async (user) => {
        const rolesResult = await pool.query<DbRole>(
          `SELECT r.* FROM "Role" r JOIN "_RoleToUser" ru ON ru."A" = r.id WHERE ru."B" = $1`,
          [user.id]
        )
        return { ...user, Role: rolesResult.rows }
      })
    )

    return usersWithRoles
  }

  /**
   * Find user by ID with roles, subjects, and classes
   */
  async findById(id: string): Promise<DbUserWithRelations> {
    const userResult = await pool.query<DbUser>(
      `SELECT * FROM "User" WHERE id = $1`,
      [id]
    )

    if (userResult.rows.length === 0) {
      throw new NotFoundError(`User with ID ${id} not found`)
    }

    const user = userResult.rows[0]

    const [rolesResult, subjectsResult, classesResult] = await Promise.all([
      pool.query<DbRole>(
        `SELECT r.* FROM "Role" r JOIN "_RoleToUser" ru ON ru."A" = r.id WHERE ru."B" = $1`,
        [id]
      ),
      pool.query<DbSubject>(
        `SELECT s.* FROM "Subject" s JOIN "_SubjectToUser" su ON su."A" = s.id WHERE su."B" = $1`,
        [id]
      ),
      pool.query<DbClass>(
        `SELECT c.* FROM "Class" c JOIN "_ClassToUser" cu ON cu."A" = c.id WHERE cu."B" = $1`,
        [id]
      )
    ])

    return {
      ...user,
      Role: rolesResult.rows,
      Subject: subjectsResult.rows,
      Class: classesResult.rows
    }
  }

  /**
   * Find user by username
   */
  async findByUsername(username: string): Promise<DbUserWithRoles | null> {
    const userResult = await pool.query<DbUser>(
      `SELECT * FROM "User" WHERE "username" = $1`,
      [username]
    )

    if (userResult.rows.length === 0) {
      return null
    }

    const user = userResult.rows[0]

    const rolesResult = await pool.query<DbRole>(
      `SELECT r.* FROM "Role" r JOIN "_RoleToUser" ru ON ru."A" = r.id WHERE ru."B" = $1`,
      [user.id]
    )

    return { ...user, Role: rolesResult.rows }
  }

  /**
   * Create a new user
   */
  async create(data: {
    username: string
    password: string
    roleNames?: string[]
  }): Promise<DbUserWithRoles> {
    const { roleNames = [], username, password } = data
    const client = await pool.connect()

    try {
      await client.query('BEGIN')

      const userId = createId()

      await client.query(
        `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)`,
        [userId, username, password]
      )

      const roles: DbRole[] = []

      for (const roleName of roleNames) {
        await client.query(
          `INSERT INTO "Role" (id, name) VALUES ($1, $2) ON CONFLICT (name) DO NOTHING`,
          [createId(), roleName]
        )
        const roleResult = await client.query<DbRole>(
          `SELECT id, name FROM "Role" WHERE name = $1`,
          [roleName]
        )
        const role = roleResult.rows[0]
        roles.push(role)

        await client.query(
          `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [role.id, userId]
        )
      }

      await client.query('COMMIT')

      const userResult = await pool.query<DbUser>(
        `SELECT * FROM "User" WHERE id = $1`,
        [userId]
      )

      return { ...userResult.rows[0], Role: roles }
    } catch (err) {
      await client.query('ROLLBACK')
      throw err
    } finally {
      client.release()
    }
  }

  /**
   * Update user roles
   */
  async updateRoles(userId: string, roleNames: string[]): Promise<DbUserWithRoles> {
    const client = await pool.connect()
    const roles: DbRole[] = []

    try {
      await client.query('BEGIN')

      await client.query(
        `DELETE FROM "_RoleToUser" WHERE "B" = $1`,
        [userId]
      )

      for (const roleName of roleNames) {
        await client.query(
          `INSERT INTO "Role" (id, name) VALUES ($1, $2) ON CONFLICT (name) DO NOTHING`,
          [createId(), roleName]
        )
        const roleResult = await client.query<DbRole>(
          `SELECT id, name FROM "Role" WHERE name = $1`,
          [roleName]
        )
        const role = roleResult.rows[0]
        roles.push(role)

        await client.query(
          `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [role.id, userId]
        )
      }

      await client.query('COMMIT')
    } catch (err) {
      await client.query('ROLLBACK')
      throw err
    } finally {
      client.release()
    }

    const { rows } = await pool.query<DbUser>('SELECT * FROM "User" WHERE id = $1', [userId])
    if (rows.length === 0) throw new NotFoundError(`User with ID ${userId} not found`)
    return { ...rows[0], Role: roles }
  }

  /**
   * Delete a user
   */
  async delete(userId: string): Promise<DbUser> {
    const { rows } = await pool.query<DbUser>(
      `DELETE FROM "User" WHERE id = $1 RETURNING *`,
      [userId]
    )
    if (rows.length === 0) {
      throw new NotFoundError(`User with ID ${userId} not found`)
    }
    return rows[0]
  }
}

export const userRepository = new UserRepository()
