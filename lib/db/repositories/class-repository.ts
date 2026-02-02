import { prisma } from '@/lib/prisma'
import { decrypt } from '@/lib/encryption'
import { NotFoundError, ForbiddenError } from '@/lib/errors'
import { createId } from '@paralleldrive/cuid2'

/**
 * Repository for Class data access
 */
export class ClassRepository {
  /**
   * Find class by ID with basic info
   */
  async findById(id: string) {
    const cls = await prisma.class.findUnique({
      where: { id },
      include: {
        Subject: true,
        User: {
          select: {
            id: true,
            username: true
          }
        }
      }
    })

    if (!cls) {
      throw new NotFoundError(`Class with ID ${id} not found`)
    }

    return cls
  }

  /**
   * Find class by ID with full assignment details
   * Automatically decrypts pupil PII
   */
  async findByIdWithAssignments(id: string) {
    const cls = await prisma.class.findUnique({
      where: { id },
      include: {
        User: {
          select: { id: true, username: true }
        },
        Subject: {
          include: {
            CommentGroup: {
              orderBy: { displayOrder: 'asc' },
              include: {
                CommentOption: {
                  orderBy: { displayOrder: 'asc' }
                }
              }
            }
          }
        },
        Assignment: {
          where: {
            Pupil: { isActive: true }
          },
          include: {
            Pupil: true,
            PupilCode: {
              include: {
                CommentGroup: true
              }
            }
          },
          orderBy: {
            Pupil: {
              lastName: 'asc'
            }
          }
        }
      }
    })

    if (!cls) {
      throw new NotFoundError(`Class with ID ${id} not found`)
    }

    // Decrypt pupil PII
    const decryptedAssignments = cls.Assignment.map(assignment => ({
      ...assignment,
      Pupil: {
        ...assignment.Pupil,
        firstName: decrypt(assignment.Pupil.firstName),
        lastName: decrypt(assignment.Pupil.lastName)
      }
    }))

    return {
      ...cls,
      Assignment: decryptedAssignments
    }
  }

  /**
   * Find all classes for a specific teacher
   */
  async findByTeacherId(teacherId: string) {
    return prisma.class.findMany({
      where: {
        User: {
          some: {
            id: teacherId
          }
        }
      },
      include: {
        Subject: true,
        _count: {
          select: {
            Assignment: true
          }
        }
      },
      orderBy: { name: 'asc' }
    })
  }

  /**
   * Assign teachers to a class
   */
  async assignTeachers(classId: string, teacherIds: string[]) {
    return prisma.class.update({
      where: { id: classId },
      data: {
        User: {
          set: teacherIds.map(id => ({ id }))
        }
      },
      include: {
        User: true
      }
    })
  }

  /**
   * Check if a user has access to a class
   * Admins and HODs of the subject have access
   * Teachers must be assigned to the class
   */
  async authorizeAccess(classId: string, userId: string, userRoles: string[]): Promise<boolean> {
    // Admins have access to everything
    if (userRoles.includes('admin')) {
      return true
    }

    const cls = await this.findById(classId)

    // HODs have access to classes in their subjects
    if (userRoles.includes('hod')) {
      const isHodOfSubject = await prisma.subject.findFirst({
        where: {
          id: cls.subjectId,
          User: {
            some: { id: userId }
          }
        }
      })
      if (isHodOfSubject) {
        return true
      }
    }

    // Teachers must be assigned to the class
    const isTeacher = cls.User.some(teacher => teacher.id === userId)
    return isTeacher
  }

  /**
   * Require access to a class (throws if unauthorized)
   */
  async requireAccess(classId: string, userId: string, userRoles: string[]): Promise<void> {
    const hasAccess = await this.authorizeAccess(classId, userId, userRoles)
    if (!hasAccess) {
      throw new ForbiddenError('You do not have access to this class')
    }
  }

  /**
   * Find all classes with counts
   */
  async findAll() {
    return prisma.class.findMany({
      include: {
        Subject: {
          select: {
            code: true,
            title: true
          }
        },
        _count: {
          select: {
            Assignment: true
          }
        },
        User: {
          select: {
            id: true,
            username: true
          }
        }
      },
      orderBy: { name: 'asc' }
    })
  }

  /**
   * Create a new class
   */
  async create(data: { name: string; year: string | null; subjectId: string }) {
    return prisma.class.create({
      data: {
        id: createId(),
        name: data.name,
        year: data.year,
        subjectId: data.subjectId
      }
    })
  }

  /**
   * Update a class
   */
  async update(classId: string, data: { name?: string; year?: string | null; subjectId?: string }) {
    return prisma.class.update({
      where: { id: classId },
      data,
      include: {
        Subject: {
          select: {
            code: true,
            title: true
          }
        },
        User: {
          select: {
            id: true,
            username: true
          }
        }
      }
    })
  }

  /**
   * Delete a class
   */
  async delete(classId: string) {
    return prisma.class.delete({
      where: { id: classId }
    })
  }

  /**
   * Assign pupils to a class by creating assignments
   */
  async assignPupils(classId: string, pupilAdmissionNumbers: string[]) {
    const assignments = pupilAdmissionNumbers.map(admissionNumber => ({
      id: `${classId}-${admissionNumber}`,
      classId,
      pupilId: admissionNumber
    }))

    return prisma.assignment.createMany({
      data: assignments,
      skipDuplicates: true
    })
  }

  /**
   * Get pupils assigned to a class
   * Automatically decrypts PII
   */
  async getPupils(classId: string) {
    const assignments = await prisma.assignment.findMany({
      where: { classId },
      include: {
        Pupil: true
      },
      orderBy: {
        Pupil: { lastName: 'asc' }
      }
    })

    return assignments.map(a => ({
      ...a.Pupil,
      firstName: decrypt(a.Pupil.firstName),
      lastName: decrypt(a.Pupil.lastName),
      assignmentId: a.id
    }))
  }

  /**
   * Remove a pupil from a class by deleting the assignment
   */
  async removePupil(classId: string, pupilAdmissionNumber: string) {
    return prisma.assignment.deleteMany({
      where: {
        classId,
        pupilId: pupilAdmissionNumber
      }
    })
  }

  /**
   * Remove multiple pupils from a class
   */
  async removePupils(classId: string, pupilAdmissionNumbers: string[]) {
    return prisma.assignment.deleteMany({
      where: {
        classId,
        pupilId: { in: pupilAdmissionNumbers }
      }
    })
  }
}


export const classRepository = new ClassRepository()
