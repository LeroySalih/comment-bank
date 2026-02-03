"use client"

import { useState, useEffect, useCallback } from "react"
import { getAuditLogs, getAuditLogActions, getAuditLogEntityTypes, getAuditLogStats, type AuditLogEntry } from "@/lib/server-actions/audit-log"

// Action display names and categories
const ACTION_LABELS: Record<string, { label: string; category: string; color: string }> = {
  // Authentication
  sign_in: { label: "Sign In", category: "Authentication", color: "bg-green-100 text-green-800" },
  sign_out: { label: "Sign Out", category: "Authentication", color: "bg-gray-100 text-gray-800" },
  sign_in_failed: { label: "Failed Sign In", category: "Authentication", color: "bg-red-100 text-red-800" },
  // User management
  create_user: { label: "Create User", category: "User", color: "bg-blue-100 text-blue-800" },
  update_user_roles: { label: "Update User Roles", category: "User", color: "bg-blue-100 text-blue-800" },
  delete_user: { label: "Delete User", category: "User", color: "bg-red-100 text-red-800" },
  // Pupil management
  create_pupil: { label: "Create Pupil", category: "Pupil", color: "bg-purple-100 text-purple-800" },
  update_pupil: { label: "Update Pupil", category: "Pupil", color: "bg-purple-100 text-purple-800" },
  delete_pupil: { label: "Delete Pupil", category: "Pupil", color: "bg-red-100 text-red-800" },
  upload_pupils: { label: "Upload Pupils", category: "Pupil", color: "bg-purple-100 text-purple-800" },
  // Subject management
  create_subject: { label: "Create Subject", category: "Subject", color: "bg-yellow-100 text-yellow-800" },
  update_subject: { label: "Update Subject", category: "Subject", color: "bg-yellow-100 text-yellow-800" },
  delete_subject: { label: "Delete Subject", category: "Subject", color: "bg-red-100 text-red-800" },
  assign_user_to_subject: { label: "Assign HOD", category: "Subject", color: "bg-yellow-100 text-yellow-800" },
  remove_user_from_subject: { label: "Remove HOD", category: "Subject", color: "bg-yellow-100 text-yellow-800" },
  // Class management
  create_class: { label: "Create Class", category: "Class", color: "bg-indigo-100 text-indigo-800" },
  update_class: { label: "Update Class", category: "Class", color: "bg-indigo-100 text-indigo-800" },
  delete_class: { label: "Delete Class", category: "Class", color: "bg-red-100 text-red-800" },
  assign_teachers_to_class: { label: "Assign Teachers", category: "Class", color: "bg-indigo-100 text-indigo-800" },
  add_pupils_to_class: { label: "Add Pupils", category: "Class", color: "bg-indigo-100 text-indigo-800" },
  remove_pupils_from_class: { label: "Remove Pupils", category: "Class", color: "bg-indigo-100 text-indigo-800" },
  // Comment management
  create_comment_group: { label: "Create Comment Group", category: "Comment", color: "bg-teal-100 text-teal-800" },
  update_comment_group: { label: "Update Comment Group", category: "Comment", color: "bg-teal-100 text-teal-800" },
  delete_comment_group: { label: "Delete Comment Group", category: "Comment", color: "bg-red-100 text-red-800" },
  reorder_comment_groups: { label: "Reorder Groups", category: "Comment", color: "bg-teal-100 text-teal-800" },
  create_comment_option: { label: "Create Comment", category: "Comment", color: "bg-teal-100 text-teal-800" },
  update_comment_option: { label: "Update Comment", category: "Comment", color: "bg-teal-100 text-teal-800" },
  delete_comment_option: { label: "Delete Comment", category: "Comment", color: "bg-red-100 text-red-800" },
  reorder_comment_options: { label: "Reorder Comments", category: "Comment", color: "bg-teal-100 text-teal-800" },
  // Assignment/Report management
  create_assignment: { label: "Create Assignment", category: "Report", color: "bg-orange-100 text-orange-800" },
  update_assignment: { label: "Update Assignment", category: "Report", color: "bg-orange-100 text-orange-800" },
  delete_assignment: { label: "Delete Assignment", category: "Report", color: "bg-red-100 text-red-800" },
  update_assignment_code: { label: "Select Comment Code", category: "Report", color: "bg-orange-100 text-orange-800" },
  update_assignment_comment: { label: "Edit Report Comment", category: "Report", color: "bg-orange-100 text-orange-800" },
  revert_assignment_comment: { label: "Revert Comment", category: "Report", color: "bg-orange-100 text-orange-800" },
  // Review
  review_comment_approved: { label: "Approve Comment", category: "Review", color: "bg-green-100 text-green-800" },
  review_comment_rejected: { label: "Reject Comment", category: "Review", color: "bg-red-100 text-red-800" },
  reset_comment_status: { label: "Resubmit Comment", category: "Review", color: "bg-yellow-100 text-yellow-800" },
  // Deadline management
  create_deadline: { label: "Create Deadline", category: "Deadline", color: "bg-pink-100 text-pink-800" },
  update_deadline: { label: "Update Deadline", category: "Deadline", color: "bg-pink-100 text-pink-800" },
  delete_deadline: { label: "Delete Deadline", category: "Deadline", color: "bg-red-100 text-red-800" },
}

function getActionInfo(action: string) {
  return ACTION_LABELS[action] || { label: action, category: "Other", color: "bg-gray-100 text-gray-800" }
}

function formatDate(date: Date) {
  return new Date(date).toLocaleString('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  })
}

function formatChanges(details: Record<string, any> | null): string {
  if (!details) return ""

  const parts: string[] = []

  if (details.before && details.after) {
    const beforeKeys = Object.keys(details.before)
    const afterKeys = Object.keys(details.after)
    const allKeys = [...new Set([...beforeKeys, ...afterKeys])]

    for (const key of allKeys) {
      if (key === 'groupId') continue // Skip internal IDs
      const before = details.before[key]
      const after = details.after[key]
      if (JSON.stringify(before) !== JSON.stringify(after)) {
        if (before !== undefined && after !== undefined) {
          parts.push(`${key}: "${before}" -> "${after}"`)
        } else if (after !== undefined) {
          parts.push(`${key}: "${after}"`)
        }
      }
    }
  } else if (details.after) {
    for (const [key, value] of Object.entries(details.after)) {
      if (typeof value === 'string' || typeof value === 'number') {
        parts.push(`${key}: "${value}"`)
      }
    }
  } else if (details.before) {
    for (const [key, value] of Object.entries(details.before)) {
      if (typeof value === 'string' || typeof value === 'number') {
        parts.push(`${key}: "${value}"`)
      }
    }
  }

  // Add other relevant details
  if (details.username) parts.push(`user: ${details.username}`)
  if (details.className) parts.push(`class: ${details.className}`)
  if (details.subjectCode) parts.push(`subject: ${details.subjectCode}`)
  if (details.note) parts.push(`note: "${details.note}"`)
  if (details.reason) parts.push(`reason: ${details.reason}`)
  if (details.pupilCount !== undefined) parts.push(`pupils: ${details.pupilCount}`)
  if (details.totalInFile !== undefined) parts.push(`uploaded: ${details.created}/${details.totalInFile}`)

  return parts.join(", ")
}

interface StatsData {
  totalLogs: number
  todayLogs: number
  weekLogs: number
  recentSignIns: number
  recentChanges: number
}

export function ActivityLog() {
  const [logs, setLogs] = useState<AuditLogEntry[]>([])
  const [total, setTotal] = useState(0)
  const [page, setPage] = useState(1)
  const [pageSize] = useState(25)
  const [totalPages, setTotalPages] = useState(1)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  // Filters
  const [actionFilter, setActionFilter] = useState<string>("")
  const [entityTypeFilter, setEntityTypeFilter] = useState<string>("")
  const [availableActions, setAvailableActions] = useState<string[]>([])
  const [availableEntityTypes, setAvailableEntityTypes] = useState<string[]>([])

  // Stats
  const [stats, setStats] = useState<StatsData | null>(null)

  // Expanded rows
  const [expandedRows, setExpandedRows] = useState<Set<string>>(new Set())

  const loadLogs = useCallback(async () => {
    setLoading(true)
    setError(null)

    try {
      const result = await getAuditLogs(page, pageSize, {
        action: actionFilter || undefined,
        entityType: entityTypeFilter || undefined
      })

      if (result.success) {
        setLogs(result.data.logs)
        setTotal(result.data.total)
        setTotalPages(result.data.totalPages)
      } else {
        setError(result.error)
      }
    } catch (e) {
      setError("Failed to load activity logs")
    } finally {
      setLoading(false)
    }
  }, [page, pageSize, actionFilter, entityTypeFilter])

  const loadFilters = useCallback(async () => {
    try {
      const [actionsResult, entityTypesResult, statsResult] = await Promise.all([
        getAuditLogActions(),
        getAuditLogEntityTypes(),
        getAuditLogStats()
      ])

      if (actionsResult.success) {
        setAvailableActions(actionsResult.actions)
      }
      if (entityTypesResult.success) {
        setAvailableEntityTypes(entityTypesResult.entityTypes)
      }
      if (statsResult.success) {
        setStats(statsResult.stats)
      }
    } catch (e) {
      console.error("Failed to load filters:", e)
    }
  }, [])

  useEffect(() => {
    loadFilters()
  }, [loadFilters])

  useEffect(() => {
    loadLogs()
  }, [loadLogs])

  const toggleRow = (id: string) => {
    const newExpanded = new Set(expandedRows)
    if (newExpanded.has(id)) {
      newExpanded.delete(id)
    } else {
      newExpanded.add(id)
    }
    setExpandedRows(newExpanded)
  }

  const handleFilterChange = () => {
    setPage(1) // Reset to first page when filters change
  }

  return (
    <div className="space-y-6">
      {/* Stats Cards */}
      {stats && (
        <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
          <div className="bg-white rounded-lg border p-4">
            <div className="text-2xl font-bold text-gray-900">{stats.totalLogs.toLocaleString()}</div>
            <div className="text-sm text-gray-500">Total Events</div>
          </div>
          <div className="bg-white rounded-lg border p-4">
            <div className="text-2xl font-bold text-blue-600">{stats.todayLogs.toLocaleString()}</div>
            <div className="text-sm text-gray-500">Today</div>
          </div>
          <div className="bg-white rounded-lg border p-4">
            <div className="text-2xl font-bold text-purple-600">{stats.weekLogs.toLocaleString()}</div>
            <div className="text-sm text-gray-500">This Week</div>
          </div>
          <div className="bg-white rounded-lg border p-4">
            <div className="text-2xl font-bold text-green-600">{stats.recentSignIns.toLocaleString()}</div>
            <div className="text-sm text-gray-500">Sign Ins Today</div>
          </div>
          <div className="bg-white rounded-lg border p-4">
            <div className="text-2xl font-bold text-orange-600">{stats.recentChanges.toLocaleString()}</div>
            <div className="text-sm text-gray-500">Changes Today</div>
          </div>
        </div>
      )}

      {/* Filters */}
      <div className="bg-white rounded-lg border p-4">
        <div className="flex flex-wrap gap-4 items-end">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Action</label>
            <select
              className="border rounded-md px-3 py-2 text-sm min-w-[200px]"
              value={actionFilter}
              onChange={(e) => {
                setActionFilter(e.target.value)
                handleFilterChange()
              }}
            >
              <option value="">All Actions</option>
              {availableActions.map(action => (
                <option key={action} value={action}>
                  {getActionInfo(action).label}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Entity Type</label>
            <select
              className="border rounded-md px-3 py-2 text-sm min-w-[150px]"
              value={entityTypeFilter}
              onChange={(e) => {
                setEntityTypeFilter(e.target.value)
                handleFilterChange()
              }}
            >
              <option value="">All Types</option>
              {availableEntityTypes.map(type => (
                <option key={type} value={type}>
                  {type.charAt(0).toUpperCase() + type.slice(1).replace('_', ' ')}
                </option>
              ))}
            </select>
          </div>
          <button
            onClick={() => {
              setActionFilter("")
              setEntityTypeFilter("")
              handleFilterChange()
            }}
            className="px-3 py-2 text-sm text-gray-600 hover:text-gray-900"
          >
            Clear Filters
          </button>
          <div className="ml-auto text-sm text-gray-500">
            Showing {logs.length} of {total.toLocaleString()} events
          </div>
        </div>
      </div>

      {/* Error State */}
      {error && (
        <div className="bg-red-50 border border-red-200 rounded-lg p-4 text-red-700">
          {error}
        </div>
      )}

      {/* Loading State */}
      {loading && (
        <div className="text-center py-8 text-gray-500">
          Loading activity logs...
        </div>
      )}

      {/* Logs Table */}
      {!loading && !error && (
        <div className="bg-white rounded-lg border overflow-hidden">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Time
                </th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  User
                </th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Action
                </th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                  Details
                </th>
                <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-8">
                </th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {logs.length === 0 ? (
                <tr>
                  <td colSpan={5} className="px-4 py-8 text-center text-gray-500">
                    No activity logs found
                  </td>
                </tr>
              ) : (
                logs.map((log) => {
                  const actionInfo = getActionInfo(log.action)
                  const isExpanded = expandedRows.has(log.id)
                  const changesSummary = formatChanges(log.details)

                  return (
                    <>
                      <tr
                        key={log.id}
                        className="hover:bg-gray-50 cursor-pointer"
                        onClick={() => toggleRow(log.id)}
                      >
                        <td className="px-4 py-3 whitespace-nowrap text-sm text-gray-500">
                          {formatDate(log.createdAt)}
                        </td>
                        <td className="px-4 py-3 whitespace-nowrap text-sm font-medium text-gray-900">
                          {log.username || "System"}
                        </td>
                        <td className="px-4 py-3 whitespace-nowrap">
                          <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${actionInfo.color}`}>
                            {actionInfo.label}
                          </span>
                        </td>
                        <td className="px-4 py-3 text-sm text-gray-500 max-w-md truncate">
                          {changesSummary || "-"}
                        </td>
                        <td className="px-4 py-3 text-sm text-gray-400">
                          {log.details && (
                            <span className="text-xs">
                              {isExpanded ? "▼" : "▶"}
                            </span>
                          )}
                        </td>
                      </tr>
                      {isExpanded && log.details && (
                        <tr key={`${log.id}-details`}>
                          <td colSpan={5} className="px-4 py-3 bg-gray-50">
                            <div className="text-sm">
                              <div className="font-medium text-gray-700 mb-2">Full Details:</div>
                              <pre className="bg-gray-100 rounded p-3 text-xs overflow-x-auto">
                                {JSON.stringify(log.details, null, 2)}
                              </pre>
                              {log.ipAddress && (
                                <div className="mt-2 text-gray-500">
                                  IP: {log.ipAddress}
                                </div>
                              )}
                            </div>
                          </td>
                        </tr>
                      )}
                    </>
                  )
                })
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* Pagination */}
      {totalPages > 1 && (
        <div className="flex items-center justify-between">
          <button
            onClick={() => setPage(p => Math.max(1, p - 1))}
            disabled={page === 1}
            className="px-4 py-2 border rounded-md text-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-gray-50"
          >
            Previous
          </button>
          <span className="text-sm text-gray-600">
            Page {page} of {totalPages}
          </span>
          <button
            onClick={() => setPage(p => Math.min(totalPages, p + 1))}
            disabled={page === totalPages}
            className="px-4 py-2 border rounded-md text-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-gray-50"
          >
            Next
          </button>
        </div>
      )}
    </div>
  )
}
