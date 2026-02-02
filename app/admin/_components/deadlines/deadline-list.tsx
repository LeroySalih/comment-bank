"use client"

import { EditDeadlineForm } from "./edit-deadline-form"

interface Deadline {
  id: string
  title: string
  date: Date | string
  description: string | null
  isActive: boolean
}

interface DeadlineListProps {
  deadlines: Deadline[]
}

export function DeadlineList({ deadlines }: DeadlineListProps) {
  const formatDate = (date: string | Date) => {
    return new Date(date).toLocaleDateString('en-GB', {
      weekday: 'short',
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    })
  }

  const isUpcoming = (date: string | Date) => new Date(date) >= new Date()
  const isPast = (date: string | Date) => new Date(date) < new Date()

  return (
    <div className="space-y-4">
      <h3 className="text-lg font-semibold text-gray-900">All Deadlines</h3>

      {deadlines.length === 0 ? (
        <div className="text-center py-10 bg-gray-50 rounded-lg border-2 border-dashed border-gray-200">
          <svg className="mx-auto h-12 w-12 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
          </svg>
          <p className="mt-2 text-gray-500">No deadlines set. Add one to get started.</p>
        </div>
      ) : (
        <div className="grid gap-4">
          {deadlines.map((deadline) => (
            <div
              key={deadline.id}
              className={`relative group bg-white shadow rounded-lg p-4 border-l-4 ${
                !deadline.isActive
                  ? 'border-gray-300 opacity-60'
                  : isPast(deadline.date)
                  ? 'border-gray-400'
                  : 'border-red-500'
              }`}
            >
              <div className="flex justify-between items-start">
                <div className="flex-1">
                  <div className="flex items-center gap-2">
                    <h4 className="font-semibold text-gray-900">{deadline.title}</h4>
                    {!deadline.isActive && (
                      <span className="text-xs bg-gray-200 text-gray-600 px-2 py-0.5 rounded">Inactive</span>
                    )}
                    {isPast(deadline.date) && deadline.isActive && (
                      <span className="text-xs bg-gray-200 text-gray-600 px-2 py-0.5 rounded">Past</span>
                    )}
                  </div>
                  <p className={`text-sm mt-1 ${isUpcoming(deadline.date) && deadline.isActive ? 'text-red-600 font-medium' : 'text-gray-500'}`}>
                    {formatDate(deadline.date)}
                  </p>
                  {deadline.description && (
                    <p className="text-sm text-gray-600 mt-2">{deadline.description}</p>
                  )}
                </div>
                <EditDeadlineForm deadline={deadline} />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}
