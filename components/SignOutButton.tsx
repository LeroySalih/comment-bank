'use client'

import { signOut } from "next-auth/react"
import { LogOut } from "lucide-react"

export default function SignOutButton() {
  return (
    <button 
      onClick={() => signOut({ callbackUrl: '/login' })}
      className="flex items-center gap-2 px-3 py-2 text-sm font-medium text-gray-600 hover:text-gray-900 bg-white hover:bg-gray-50 border border-gray-200 rounded-md transition-colors shadow-sm"
    >
      <LogOut className="w-4 h-4" />
      Sign Out
    </button>
  )
}
