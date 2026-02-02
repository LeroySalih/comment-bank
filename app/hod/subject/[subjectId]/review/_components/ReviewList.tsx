'use client'

import { useState } from 'react'
import { useRouter } from 'next/navigation'
import Link from 'next/link'
import CommentStatusBadge from '@/components/CommentStatusBadge'
import { reviewComment } from '@/lib/server-actions/comment-check'

type Assignment = {
  id: string
  finalComment?: string | null
  checkStatus: string
  checkNote?: string | null
  checkedAt?: Date | null
  Pupil: {
    firstName: string
    lastName: string
    gender: string
  }
  Class: {
    id: string
    name: string
  }
  CheckedBy?: {
    username: string
  } | null
}

interface ReviewListProps {
  assignments: Assignment[]
  subjectId: string
}

export default function ReviewList({ assignments, subjectId }: ReviewListProps) {
  const [expandedId, setExpandedId] = useState<string | null>(null)
  const [reviewingId, setReviewingId] = useState<string | null>(null)
  const [note, setNote] = useState('')
  const [submitting, setSubmitting] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const router = useRouter()

  const handleApprove = async (assignmentId: string) => {
    setSubmitting(true)
    setError(null)

    const result = await reviewComment(assignmentId, 'checked_ok')

    if (result.success) {
      router.refresh()
      setReviewingId(null)
    } else {
      setError(result.error || 'Failed to approve comment')
    }

    setSubmitting(false)
  }

  const handleReject = async (assignmentId: string) => {
    if (!note.trim()) {
      setError('Please provide feedback for the rejection')
      return
    }

    setSubmitting(true)
    setError(null)

    const result = await reviewComment(assignmentId, 'checked_rejected', note)

    if (result.success) {
      router.refresh()
      setReviewingId(null)
      setNote('')
    } else {
      setError(result.error || 'Failed to reject comment')
    }

    setSubmitting(false)
  }

  return (
    <div className="space-y-4">
      {assignments.map(assignment => (
        <div
          key={assignment.id}
          className="bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 overflow-hidden"
        >
          {/* Header Row */}
          <div
            className="px-6 py-4 flex items-center justify-between cursor-pointer hover:bg-gray-50 dark:hover:bg-gray-800/50"
            onClick={() => setExpandedId(expandedId === assignment.id ? null : assignment.id)}
          >
            <div className="flex items-center gap-4">
              <div>
                <p className="text-[#111418] dark:text-white font-semibold">
                  {assignment.Pupil.lastName}, {assignment.Pupil.firstName}
                </p>
                <p className="text-sm text-gray-500">{assignment.Class.name}</p>
              </div>
            </div>
            <div className="flex items-center gap-4">
              <CommentStatusBadge status={assignment.checkStatus} size="sm" />
              <span className="material-symbols-outlined text-gray-400">
                {expandedId === assignment.id ? 'expand_less' : 'expand_more'}
              </span>
            </div>
          </div>

          {/* Expanded Content */}
          {expandedId === assignment.id && (
            <div className="border-t border-[#f0f2f4] dark:border-gray-800">
              {/* Comment Preview */}
              <div className="p-6 bg-gray-50 dark:bg-gray-800/30">
                <label className="text-xs font-bold text-gray-500 uppercase mb-2 block">Comment</label>
                <p className="text-gray-700 dark:text-gray-300 whitespace-pre-wrap">
                  {assignment.finalComment || <span className="italic text-gray-400">No comment text</span>}
                </p>
              </div>

              {/* Previous Check Note */}
              {assignment.checkNote && assignment.checkStatus === 'checked_rejected' && (
                <div className="px-6 py-4 bg-red-50 dark:bg-red-900/20 border-t border-red-100 dark:border-red-900">
                  <label className="text-xs font-bold text-red-600 uppercase mb-1 block">Rejection Feedback</label>
                  <p className="text-red-700 dark:text-red-400 text-sm">{assignment.checkNote}</p>
                </div>
              )}

              {/* Actions */}
              <div className="px-6 py-4 border-t border-[#f0f2f4] dark:border-gray-800 flex items-center justify-between">
                <Link
                  href={`/student/${assignment.id}`}
                  className="text-primary hover:text-blue-700 text-sm font-medium inline-flex items-center gap-1"
                >
                  <span className="material-symbols-outlined text-lg">visibility</span>
                  View Full Editor
                </Link>

                {assignment.checkStatus === 'required_check' ? (
                  reviewingId === assignment.id ? (
                    <div className="flex items-center gap-3">
                      <input
                        type="text"
                        value={note}
                        onChange={e => setNote(e.target.value)}
                        placeholder="Feedback for rejection (required)"
                        className="border border-gray-300 dark:border-gray-600 rounded-lg px-3 py-2 text-sm w-64 dark:bg-gray-800"
                      />
                      <button
                        onClick={() => handleReject(assignment.id)}
                        disabled={submitting}
                        className="px-4 py-2 bg-red-600 text-white text-sm font-medium rounded-lg hover:bg-red-700 disabled:opacity-50"
                      >
                        {submitting ? 'Saving...' : 'Reject'}
                      </button>
                      <button
                        onClick={() => {
                          setReviewingId(null)
                          setNote('')
                          setError(null)
                        }}
                        className="px-4 py-2 bg-gray-200 dark:bg-gray-700 text-gray-700 dark:text-gray-300 text-sm font-medium rounded-lg hover:bg-gray-300 dark:hover:bg-gray-600"
                      >
                        Cancel
                      </button>
                    </div>
                  ) : (
                    <div className="flex items-center gap-3">
                      <button
                        onClick={() => handleApprove(assignment.id)}
                        disabled={submitting}
                        className="px-4 py-2 bg-emerald-600 text-white text-sm font-medium rounded-lg hover:bg-emerald-700 disabled:opacity-50 inline-flex items-center gap-1"
                      >
                        <span className="material-symbols-outlined text-lg">check</span>
                        Approve
                      </button>
                      <button
                        onClick={() => setReviewingId(assignment.id)}
                        className="px-4 py-2 bg-red-100 text-red-700 text-sm font-medium rounded-lg hover:bg-red-200 inline-flex items-center gap-1"
                      >
                        <span className="material-symbols-outlined text-lg">close</span>
                        Reject
                      </button>
                    </div>
                  )
                ) : assignment.checkStatus === 'checked_ok' ? (
                  <span className="text-emerald-600 text-sm font-medium inline-flex items-center gap-1">
                    <span className="material-symbols-outlined">verified</span>
                    Approved {assignment.CheckedBy && `by ${assignment.CheckedBy.username}`}
                  </span>
                ) : assignment.checkStatus === 'checked_rejected' ? (
                  <span className="text-red-600 text-sm font-medium inline-flex items-center gap-1">
                    <span className="material-symbols-outlined">error</span>
                    Rejected {assignment.CheckedBy && `by ${assignment.CheckedBy.username}`}
                  </span>
                ) : null}
              </div>

              {error && (
                <div className="px-6 pb-4">
                  <p className="text-red-600 text-sm">{error}</p>
                </div>
              )}
            </div>
          )}
        </div>
      ))}
    </div>
  )
}
