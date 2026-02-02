import { cache } from 'react'
import { userRepository } from '../repositories/user-repository'

/**
 * Cached query to get all users
 * Uses React's cache() for request-level memoization
 */
export const getUsers = cache(async () => {
  return userRepository.findAll()
})

/**
 * Cached query to get a single user
 */
export const getUser = cache(async (id: string) => {
  return userRepository.findById(id)
})
