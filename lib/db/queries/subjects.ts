import { cache } from 'react'
import { subjectRepository } from '../repositories/subject-repository'

/**
 * Cached query to get all subjects
 * Uses React's cache() for request-level memoization
 */
export const getSubjects = cache(async () => {
  return subjectRepository.findAll()
})

/**
 * Cached query to get a single subject with full details
 */
export const getSubject = cache(async (id: string) => {
  return subjectRepository.findByIdWithDetails(id)
})
