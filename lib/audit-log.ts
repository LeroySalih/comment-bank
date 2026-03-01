import { pool } from '@/lib/db'
import { createId } from '@paralleldrive/cuid2'
import { getServerSession } from 'next-auth'
import { authOptions } from '@/app/api/auth/[...nextauth]/route'
import { headers } from 'next/headers'
import { encrypt } from '@/lib/encryption'

export type AuditAction =
  // Authentication
  | 'sign_in'
  | 'sign_out'
  | 'sign_in_failed'
  // User management
  | 'create_user'
  | 'update_user_roles'
  | 'update_user_active_status'
  | 'delete_user'
  // Pupil management
  | 'create_pupil'
  | 'update_pupil'
  | 'delete_pupil'
  | 'upload_pupils'
  // Subject management
  | 'create_subject'
  | 'update_subject'
  | 'delete_subject'
  | 'assign_user_to_subject'
  | 'remove_user_from_subject'
  // Class management
  | 'create_class'
  | 'update_class'
  | 'delete_class'
  | 'assign_teachers_to_class'
  | 'add_pupils_to_class'
  | 'remove_pupils_from_class'
  // Comment group management
  | 'create_comment_group'
  | 'update_comment_group'
  | 'delete_comment_group'
  | 'reorder_comment_groups'
  // Comment option management
  | 'create_comment_option'
  | 'update_comment_option'
  | 'delete_comment_option'
  | 'reorder_comment_options'
  // Common Comment Group management
  | 'create_common_comment_group'
  | 'update_common_comment_group'
  | 'delete_common_comment_group'
  | 'reorder_common_comment_groups'
  // Common Comment Option management
  | 'create_common_comment_option'
  | 'update_common_comment_option'
  | 'delete_common_comment_option'
  | 'reorder_common_comment_options'
  // Format template
  | 'update_comment_format_template'
  | 'update_subject_comment_format'
  // Assignment/Report management
  | 'create_assignment'
  | 'update_assignment'
  | 'delete_assignment'
  | 'update_assignment_code'
  | 'update_common_assignment_code'
  | 'update_assignment_comment'
  | 'revert_assignment_comment'
  // Comment review
  | 'review_comment_approved'
  | 'review_comment_rejected'
  | 'reset_comment_status'
  // Linked data
  | 'upload_linked_data'
  // Deadline management
  | 'create_deadline'
  | 'update_deadline'
  | 'delete_deadline'

export type EntityType =
  | 'user'
  | 'pupil'
  | 'subject'
  | 'class'
  | 'comment_group'
  | 'comment_option'
  | 'assignment'
  | 'common_comment_group'
  | 'common_comment_option'
  | 'app_setting'
  | 'deadline'

export interface AuditLogDetails {
  before?: Record<string, any>
  after?: Record<string, any>
  [key: string]: any
}

export interface CreateAuditLogParams {
  action: AuditAction
  entityType?: EntityType
  entityId?: string
  details?: AuditLogDetails
  userId?: string
  username?: string
}

/**
 * Create an audit log entry
 * Automatically captures user info from session if not provided
 */
export async function createAuditLog(params: CreateAuditLogParams): Promise<void> {
  try {
    let userId = params.userId
    let username = params.username
    let ipAddress: string | undefined
    let userAgent: string | undefined

    // Try to get user info from session if not provided
    if (!userId || !username) {
      try {
        const session = await getServerSession(authOptions)
        if (session?.user) {
          userId = userId || session.user.id
          username = username || session.user.username
        }
      } catch {
        // Session not available (e.g., during sign-in)
      }
    }

    // Try to get request headers
    try {
      const headersList = await headers()
      ipAddress = headersList.get('x-forwarded-for') || headersList.get('x-real-ip') || undefined
      userAgent = headersList.get('user-agent') || undefined
    } catch {
      // Headers not available
    }

    await pool.query(
      `INSERT INTO "AuditLog" (id, "userId", username, action, "entityType", "entityId", details, "ipAddress", "userAgent")
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)`,
      [
        createId(),
        userId ?? null,
        username ?? null,
        params.action,
        params.entityType ?? null,
        params.entityId ?? null,
        params.details ? encrypt(JSON.stringify(params.details)) : null,
        ipAddress ?? null,
        userAgent ?? null
      ]
    )
  } catch (error) {
    // Log but don't throw - audit logging shouldn't break the main operation
    console.error('Failed to create audit log:', error)
  }
}

/**
 * Helper to create audit log for data changes with before/after values
 */
export async function logDataChange(
  action: AuditAction,
  entityType: EntityType,
  entityId: string,
  before: Record<string, any> | null,
  after: Record<string, any> | null,
  additionalDetails?: Record<string, any>
): Promise<void> {
  await createAuditLog({
    action,
    entityType,
    entityId,
    details: {
      ...(before && { before }),
      ...(after && { after }),
      ...additionalDetails
    }
  })
}

/**
 * Helper to create audit log for authentication events
 */
export async function logAuthEvent(
  action: 'sign_in' | 'sign_out' | 'sign_in_failed',
  userId?: string,
  username?: string,
  details?: Record<string, any>
): Promise<void> {
  await createAuditLog({
    action,
    entityType: 'user',
    entityId: userId,
    userId,
    username,
    details
  })
}
