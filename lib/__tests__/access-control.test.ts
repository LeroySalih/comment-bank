import { describe, it, expect } from 'vitest'
import { isAdmin, isHoD, isTeacher } from '../access-control'

describe('Access Control Functions', () => {
  describe('isAdmin', () => {
    it('should return true for admin user', () => {
      const user = {
        id: '1',
        username: 'admin',
        roles: ['admin']
      }
      expect(isAdmin(user)).toBe(true)
    })

    it('should return true for user with multiple roles including admin', () => {
      const user = {
        id: '1',
        username: 'superuser',
        roles: ['admin', 'hod', 'teacher']
      }
      expect(isAdmin(user)).toBe(true)
    })

    it('should return false for non-admin user', () => {
      const user = {
        id: '1',
        username: 'teacher',
        roles: ['teacher']
      }
      expect(isAdmin(user)).toBe(false)
    })

    it('should return false for user with no roles', () => {
      const user = {
        id: '1',
        username: 'user',
        roles: []
      }
      expect(isAdmin(user)).toBe(false)
    })

    it('should return false for undefined user', () => {
      expect(isAdmin(undefined)).toBe(false)
    })
  })

  describe('isHoD', () => {
    it('should return true for HOD user', () => {
      const user = {
        id: '1',
        username: 'hod',
        roles: ['hod']
      }
      expect(isHoD(user)).toBe(true)
    })

    it('should return true for admin (admins are also HODs)', () => {
      const user = {
        id: '1',
        username: 'admin',
        roles: ['admin']
      }
      expect(isHoD(user)).toBe(true)
    })

    it('should return false for teacher', () => {
      const user = {
        id: '1',
        username: 'teacher',
        roles: ['teacher']
      }
      expect(isHoD(user)).toBe(false)
    })

    it('should return false for undefined user', () => {
      expect(isHoD(undefined)).toBe(false)
    })
  })

  describe('isTeacher', () => {
    it('should return true for teacher user', () => {
      const user = {
        id: '1',
        username: 'teacher',
        roles: ['teacher']
      }
      expect(isTeacher(user)).toBe(true)
    })

    it('should return true for HOD (HODs are also teachers)', () => {
      const user = {
        id: '1',
        username: 'hod',
        roles: ['hod']
      }
      expect(isTeacher(user)).toBe(true)
    })

    it('should return true for admin (admins are also teachers)', () => {
      const user = {
        id: '1',
        username: 'admin',
        roles: ['admin']
      }
      expect(isTeacher(user)).toBe(true)
    })

    it('should return false for user with no roles', () => {
      const user = {
        id: '1',
        username: 'user',
        roles: []
      }
      expect(isTeacher(user)).toBe(false)
    })

    it('should return false for undefined user', () => {
      expect(isTeacher(undefined)).toBe(false)
    })
  })
})
