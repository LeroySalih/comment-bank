import { cache } from 'react'
import { classRepository } from '../repositories/class-repository'

/**
 * Cached query to get a class with full assignment details
 * Uses React's cache() for request-level memoization
 */
export const getClass = cache(async (id: string) => {
  return classRepository.findByIdWithAssignments(id)
})

/**
 * Cached query to get classes for a teacher
 */
export const getTeacherClasses = cache(async (teacherId: string) => {
  return classRepository.findByTeacherId(teacherId)
})
