import { getServerSession } from 'next-auth'
import { redirect } from 'next/navigation'
import { authOptions } from '@/app/api/auth/[...nextauth]/route'
import { logger } from '@/lib/logger'
import type { Session } from 'next-auth'

/**
 * Require specific role(s) for page access
 * Redirects to login if not authenticated
 * Redirects to home if authenticated but lacking required role
 * 
 * @param roles - Single role or array of roles required
 * @returns Session object if authorized
 * 
 * @example
 * // In a page component
 * export default async function AdminPage() {
 *   const session = await requireRole('admin')
 *   // Only admins can reach this point
 *   return <div>Admin Dashboard</div>
 * }
 * 
 * @example
 * export default async function HODPage() {
 *   const session = await requireRole(['admin', 'hod'])
 *   // Both admins and HODs can reach this point
 *   return <div>HOD Dashboard</div>
 * }
 */
export async function requireRole(roles: string | string[]): Promise<Session> {
  const session = await getServerSession(authOptions)
  const roleArray = Array.isArray(roles) ? roles : [roles]
  
  if (!session?.user) {
    logger.info('Redirecting to login - no session', {
      requiredRoles: roleArray
    })
    redirect('/login')
  }
  
  const userRoles = session.user.roles || []
  const hasRequiredRole = roleArray.some(role => userRoles.includes(role))
  
  if (!hasRequiredRole) {
    logger.warn('Access denied - insufficient permissions', {
      userId: session.user.id,
      userRoles,
      requiredRoles: roleArray
    })
    redirect('/')
  }
  
  return session
}
