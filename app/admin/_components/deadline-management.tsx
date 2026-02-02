"use client"

import { DeadlineForm } from "./deadlines/deadline-form"
import { DeadlineList } from "./deadlines/deadline-list"

interface Deadline {
  id: string
  title: string
  date: Date | string
  description: string | null
  isActive: boolean
  createdAt: Date | string
}

interface DeadlineManagementProps {
  deadlines: Deadline[]
}

export function DeadlineManagement({ deadlines }: DeadlineManagementProps) {
  return (
    <div>
      <div className="mb-8">
        <DeadlineForm />
      </div>
      <DeadlineList deadlines={deadlines} />
    </div>
  )
}
