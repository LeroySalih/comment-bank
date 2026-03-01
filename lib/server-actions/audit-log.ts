"use server"

import { pool } from '@/lib/db'
import { withRole } from '@/lib/auth/with-role'
import { handleServerActionError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import { decrypt } from '@/lib/encryption'

export interface AuditLogEntry {
  id: string
  userId: string | null
  username: string | null
  action: string
  entityType: string | null
  entityId: string | null
  details: Record<string, any> | null
  ipAddress: string | null
  userAgent: string | null
  createdAt: Date
}

export interface AuditLogFilters {
  action?: string
  entityType?: string
  userId?: string
  startDate?: Date
  endDate?: Date
}

export interface PaginatedAuditLogs {
  logs: AuditLogEntry[]
  total: number
  page: number
  pageSize: number
  totalPages: number
}

/**
 * Get audit logs with pagination and filtering (Admin only)
 */
export const getAuditLogs = withRole('admin', async (
  page: number = 1,
  pageSize: number = 50,
  filters?: AuditLogFilters
): Promise<{ success: true; data: PaginatedAuditLogs } | { success: false; error: string; code: string }> => {
  try {
    const skip = (page - 1) * pageSize

    // Build dynamic WHERE clause from filters
    const conditions: string[] = []
    const params: any[] = []
    let paramIdx = 1

    if (filters?.action) {
      conditions.push(`action = $${paramIdx++}`)
      params.push(filters.action)
    }

    if (filters?.entityType) {
      conditions.push(`"entityType" = $${paramIdx++}`)
      params.push(filters.entityType)
    }

    if (filters?.userId) {
      conditions.push(`"userId" = $${paramIdx++}`)
      params.push(filters.userId)
    }

    if (filters?.startDate) {
      conditions.push(`"createdAt" >= $${paramIdx++}`)
      params.push(filters.startDate)
    }

    if (filters?.endDate) {
      conditions.push(`"createdAt" <= $${paramIdx++}`)
      params.push(filters.endDate)
    }

    const whereClause = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : ''

    // Fetch paginated logs
    const logsParams = [...params, pageSize, skip]
    const limitIdx = paramIdx
    const offsetIdx = paramIdx + 1
    const { rows: logs } = await pool.query(
      `SELECT * FROM "AuditLog" ${whereClause} ORDER BY "createdAt" DESC LIMIT $${limitIdx} OFFSET $${offsetIdx}`,
      logsParams
    )

    // Fetch total count using same WHERE clause
    const { rows: countRows } = await pool.query(
      `SELECT COUNT(*) as count FROM "AuditLog" ${whereClause}`,
      params
    )
    const total = Number(countRows[0].count)

    // Parse details JSON
    const parsedLogs: AuditLogEntry[] = logs.map(log => ({
      ...log,
      details: log.details ? JSON.parse(decrypt(log.details)) : null
    }))

    return {
      success: true,
      data: {
        logs: parsedLogs,
        total,
        page,
        pageSize,
        totalPages: Math.ceil(total / pageSize)
      }
    }
  } catch (error) {
    logger.error('Failed to get audit logs', { error })
    return handleServerActionError(error) as { success: false; error: string; code: string }
  }
})

/**
 * Get distinct action types for filtering (Admin only)
 */
export const getAuditLogActions = withRole('admin', async () => {
  try {
    const { rows } = await pool.query(
      `SELECT DISTINCT action FROM "AuditLog" ORDER BY action ASC`
    )

    return {
      success: true as const,
      actions: rows.map(r => r.action as string)
    }
  } catch (error) {
    logger.error('Failed to get audit log actions', { error })
    return handleServerActionError(error)
  }
})

/**
 * Get distinct entity types for filtering (Admin only)
 */
export const getAuditLogEntityTypes = withRole('admin', async () => {
  try {
    const { rows } = await pool.query(
      `SELECT DISTINCT "entityType" FROM "AuditLog" WHERE "entityType" IS NOT NULL ORDER BY "entityType" ASC`
    )

    return {
      success: true as const,
      entityTypes: rows.map(r => r.entityType as string)
    }
  } catch (error) {
    logger.error('Failed to get audit log entity types', { error })
    return handleServerActionError(error)
  }
})

/**
 * Get audit log statistics for the dashboard (Admin only)
 */
export const getAuditLogStats = withRole('admin', async () => {
  try {
    const now = new Date()
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate())
    const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000)

    const { rows } = await pool.query(
      `SELECT
        COUNT(*) FILTER (WHERE true) as total,
        COUNT(*) FILTER (WHERE "createdAt" >= $1) as today,
        COUNT(*) FILTER (WHERE "createdAt" >= $2) as week,
        COUNT(*) FILTER (WHERE action = 'sign_in' AND "createdAt" >= $1) as sign_ins,
        COUNT(*) FILTER (WHERE action NOT IN ('sign_in', 'sign_out', 'sign_in_failed') AND "createdAt" >= $1) as changes
       FROM "AuditLog"`,
      [today, weekAgo]
    )

    const stats = rows[0]

    return {
      success: true as const,
      stats: {
        totalLogs: Number(stats.total),
        todayLogs: Number(stats.today),
        weekLogs: Number(stats.week),
        recentSignIns: Number(stats.sign_ins),
        recentChanges: Number(stats.changes)
      }
    }
  } catch (error) {
    logger.error('Failed to get audit log stats', { error })
    return handleServerActionError(error)
  }
})
