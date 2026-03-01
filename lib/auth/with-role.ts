import { getServerSession } from 'next-auth'
import { authOptions } from '@/app/api/auth/[...nextauth]/route'
import { AuthError, ForbiddenError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import { pool } from '@/lib/db'

/**
 * Higher-order function to wrap server actions with role-based authorization
 *
 * @param roles - Single role or array of roles required to execute the action
 * @param handler - The server action handler function
 * @returns Wrapped handler with authorization check
 *
 * @example
 * export const deleteUser = withRole('admin', async (userId: string) => {
 *   // Only admins can execute this
 *   await pool.query('DELETE FROM "User" WHERE id = $1', [userId])
 * })
 *
 * @example
 * export const updateSubject = withRole(['admin', 'hod'], async (subjectId: string, data: any) => {
 *   // Both admins and HODs can execute this
 *   await pool.query('UPDATE "Subject" SET ... WHERE id = $1', [subjectId])
 * })
 */
export function withRole<T extends any[], R>(
  roles: string | string[],
  handler: (...args: T) => Promise<R>
) {
  return async (...args: T): Promise<R> => {
    const session = await getServerSession(authOptions)
    const roleArray = Array.isArray(roles) ? roles : [roles]

    if (!session?.user) {
      logger.warn('Unauthorized access attempt - no session', {
        requiredRoles: roleArray
      })
      throw new AuthError('Not authenticated')
    }

    // Check if user is still active in the database
    const { rows } = await pool.query<{ isActive: boolean }>(
      `SELECT "isActive" FROM "User" WHERE id = $1`,
      [session.user.id]
    )
    const user = rows[0] ?? null

    if (!user || !user.isActive) {
      logger.warn('Forbidden access attempt - user is inactive', {
        userId: session.user.id
      })
      throw new ForbiddenError('Account is inactive')
    }

    const userRoles = session.user.roles || []
    const hasRequiredRole = roleArray.some(role => userRoles.includes(role))

    if (!hasRequiredRole) {
      logger.warn('Forbidden access attempt - insufficient permissions', {
        userId: session.user.id,
        userRoles,
        requiredRoles: roleArray
      })
      throw new ForbiddenError('Insufficient permissions')
    }

    logger.debug('Authorization successful', {
      userId: session.user.id,
      requiredRoles: roleArray
    })

    return handler(...args)
  }
}
