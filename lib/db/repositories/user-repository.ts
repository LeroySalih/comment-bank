import { prisma } from '@/lib/prisma'
import { NotFoundError } from '@/lib/errors'
import { createId } from '@paralleldrive/cuid2'
import type { User } from '@prisma/client'

/**
 * Repository for User data access
 */
export class UserRepository {
  /**
   * Find all users with their roles
   */
  async findAll() {
    return prisma.user.findMany({
      include: {
        Role: true
      },
      orderBy: { username: 'asc' }
    })
  }

  /**
   * Find user by ID with roles
   */
  async findById(id: string) {
    const user = await prisma.user.findUnique({
      where: { id },
      include: {
        Role: true,
        Subject: true,
        Class: true
      }
    })

    if (!user) {
      throw new NotFoundError(`User with ID ${id} not found`)
    }

    return user
  }

  /**
   * Find user by username
   */
  async findByUsername(username: string) {
    return prisma.user.findUnique({
      where: { username },
      include: {
        Role: true
      }
    })
  }

  /**
   * Create a new user
   */
  async create(data: {
    username: string
    password: string
    roleNames?: string[]
  }) {
    const { roleNames = [], ...userData } = data

    // Find or create roles
    const roles = await Promise.all(
      roleNames.map(name =>
        prisma.role.upsert({
          where: { name },
          create: { id: name, name },
          update: {}
        })
      )
    )

    return prisma.user.create({
      data: {
        id: createId(),
        ...userData,
        Role: {
          connect: roles.map(role => ({ id: role.id }))
        }
      },
      include: {
        Role: true
      }
    })
  }

  /**
   * Update user roles
   */
  async updateRoles(userId: string, roleNames: string[]) {
    // Find or create roles
    const roles = await Promise.all(
      roleNames.map(name =>
        prisma.role.upsert({
          where: { name },
          create: { id: name, name },
          update: {}
        })
      )
    )

    return prisma.user.update({
      where: { id: userId },
      data: {
        Role: {
          set: roles.map(role => ({ id: role.id }))
        }
      },
      include: {
        Role: true
      }
    })
  }

  /**
   * Delete a user
   */
  async delete(userId: string) {
    return prisma.user.delete({
      where: { id: userId }
    })
  }
}

export const userRepository = new UserRepository()
