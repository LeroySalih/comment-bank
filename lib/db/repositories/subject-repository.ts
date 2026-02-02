import { prisma } from '@/lib/prisma'
import { NotFoundError } from '@/lib/errors'

/**
 * Repository for Subject data access
 */
export class SubjectRepository {
  /**
   * Find all subjects with basic relations
   */
  async findAll() {
    return prisma.subject.findMany({
      include: {
        Class: {
          orderBy: { name: 'asc' }
        },
        User: {
          select: {
            id: true,
            username: true
          }
        },
        _count: {
          select: {
            CommentGroup: true
          }
        }
      },
      orderBy: { code: 'asc' }
    })
  }

  /**
   * Find subject by ID with full details
   */
  async findByIdWithDetails(id: string) {
    const subject = await prisma.subject.findUnique({
      where: { id },
      include: {
        Class: {
          orderBy: { name: 'asc' }
        },
        User: {
          select: {
            id: true,
            username: true
          }
        },
        CommentGroup: {
          orderBy: { displayOrder: 'asc' },
          include: {
            CommentOption: {
              orderBy: { displayOrder: 'asc' }
            }
          }
        }
      }
    })

    if (!subject) {
      throw new NotFoundError(`Subject with ID ${id} not found`)
    }

    return subject
  }

  /**
   * Find subject by code
   */
  async findByCode(code: string) {
    return prisma.subject.findUnique({
      where: { code }
    })
  }

  /**
   * Create a new subject
   */
  async create(data: {
    code: string
    title?: string
    studiedComment?: string
  }) {
    return prisma.subject.create({
      data: {
        id: data.code,
        ...data
      },
      include: {
        Class: true,
        User: true
      }
    })
  }

  /**
   * Update a subject
   */
  async update(
    id: string,
    data: {
      code?: string
      title?: string
      studiedComment?: string
    }
  ) {
    return prisma.subject.update({
      where: { id },
      data,
      include: {
        Class: true,
        User: true
      }
    })
  }

  /**
   * Delete a subject
   */
  async delete(id: string) {
    return prisma.subject.delete({
      where: { id }
    })
  }

  /**
   * Assign a user (HOD) to a subject
   */
  async assignUser(subjectId: string, userId: string) {
    return prisma.subject.update({
      where: { id: subjectId },
      data: {
        User: {
          connect: { id: userId }
        }
      },
      include: {
        User: true
      }
    })
  }

  /**
   * Remove a user from a subject
   */
  async removeUser(subjectId: string, userId: string) {
    return prisma.subject.update({
      where: { id: subjectId },
      data: {
        User: {
          disconnect: { id: userId }
        }
      },
      include: {
        User: true
      }
    })
  }
}

export const subjectRepository = new SubjectRepository()
