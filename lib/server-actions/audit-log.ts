"use server"

import { prisma } from '@/lib/prisma'
import { withRole } from '@/lib/auth/with-role'
import { handleServerActionError } from '@/lib/errors'
import { logger } from '@/lib/logger'

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

    // Build where clause from filters
    const where: any = {}

    if (filters?.action) {
      where.action = filters.action
    }

    if (filters?.entityType) {
      where.entityType = filters.entityType
    }

    if (filters?.userId) {
      where.userId = filters.userId
    }

    if (filters?.startDate || filters?.endDate) {
      where.createdAt = {}
      if (filters.startDate) {
        where.createdAt.gte = filters.startDate
      }
      if (filters.endDate) {
        where.createdAt.lte = filters.endDate
      }
    }

    const [logs, total] = await Promise.all([
      prisma.auditLog.findMany({
        where,
        orderBy: { createdAt: 'desc' },
        skip,
        take: pageSize
      }),
      prisma.auditLog.count({ where })
    ])

    // Parse details JSON
    const parsedLogs: AuditLogEntry[] = logs.map(log => ({
      ...log,
      details: log.details ? JSON.parse(log.details) : null
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
    const actions = await prisma.auditLog.groupBy({
      by: ['action'],
      orderBy: { action: 'asc' }
    })

    return {
      success: true as const,
      actions: actions.map(a => a.action)
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
    const entityTypes = await prisma.auditLog.groupBy({
      by: ['entityType'],
      where: { entityType: { not: null } },
      orderBy: { entityType: 'asc' }
    })

    return {
      success: true as const,
      entityTypes: entityTypes.map(e => e.entityType).filter(Boolean) as string[]
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

    const [totalLogs, todayLogs, weekLogs, recentSignIns, recentChanges] = await Promise.all([
      prisma.auditLog.count(),
      prisma.auditLog.count({
        where: { createdAt: { gte: today } }
      }),
      prisma.auditLog.count({
        where: { createdAt: { gte: weekAgo } }
      }),
      prisma.auditLog.count({
        where: {
          action: 'sign_in',
          createdAt: { gte: today }
        }
      }),
      prisma.auditLog.count({
        where: {
          action: { notIn: ['sign_in', 'sign_out', 'sign_in_failed'] },
          createdAt: { gte: today }
        }
      })
    ])

    return {
      success: true as const,
      stats: {
        totalLogs,
        todayLogs,
        weekLogs,
        recentSignIns,
        recentChanges
      }
    }
  } catch (error) {
    logger.error('Failed to get audit log stats', { error })
    return handleServerActionError(error)
  }
})
