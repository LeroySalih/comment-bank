import { Session } from "next-auth"

export const ROLES = {
  ADMIN: "admin",
  HOD: "hod",
  TEACHER: "teacher",
}

export function hasRole(user: { roles: string[] } | undefined | null, role: string): boolean {
  if (!user || !user.roles) return false
  return user.roles.includes(role)
}

export function isAdmin(user: { roles: string[] } | undefined | null): boolean {
  return hasRole(user, ROLES.ADMIN)
}

export function isHoD(user: { roles: string[] } | undefined | null): boolean {
  return hasRole(user, ROLES.HOD)
}

export function isTeacher(user: { roles: string[] } | undefined | null): boolean {
  return hasRole(user, ROLES.TEACHER)
}
