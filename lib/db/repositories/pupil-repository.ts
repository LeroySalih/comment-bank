import { prisma } from '@/lib/prisma'
import { encrypt, decrypt } from '@/lib/encryption'
import { NotFoundError } from '@/lib/errors'
import type { Pupil } from '@prisma/client'

/**
 * Repository for Pupil data access
 * Handles encryption/decryption automatically
 */
export class PupilRepository {
  /**
   * Find all pupils with optional search query
   * Automatically decrypts PII fields
   */
  async findAll(query?: string): Promise<Pupil[]> {
    const pupils = await prisma.pupil.findMany({
      where: query
        ? {
            OR: [
              { firstName: { contains: query } },
              { lastName: { contains: query } },
              { admissionNumber: { contains: query } }
            ]
          }
        : undefined,
      orderBy: [{ lastName: 'asc' }, { firstName: 'asc' }]
    })

    return pupils.map(pupil => ({
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }))
  }

  /**
   * Find pupil by admission number
   * Automatically decrypts PII fields
   */
  async findByAdmissionNumber(admissionNumber: string): Promise<Pupil | null> {
    const pupil = await prisma.pupil.findUnique({
      where: { admissionNumber }
    })

    if (!pupil) return null

    return {
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }
  }

  /**
   * Create a new pupil
   * Automatically encrypts PII fields
   */
  async create(data: {
    admissionNumber: string
    firstName: string
    lastName: string
    gender: string
    isActive?: boolean
  }): Promise<Pupil> {
    const pupil = await prisma.pupil.create({
      data: {
        ...data,
        firstName: encrypt(data.firstName),
        lastName: encrypt(data.lastName)
      }
    })

    return {
      ...pupil,
      firstName: data.firstName, // Return unencrypted
      lastName: data.lastName
    }
  }

  /**
   * Update a pupil
   * Automatically encrypts PII fields if provided
   */
  async update(
    admissionNumber: string,
    data: {
      firstName?: string
      lastName?: string
      gender?: string
      isActive?: boolean
    }
  ): Promise<Pupil> {
    const updateData: any = { ...data }
    
    if (data.firstName) {
      updateData.firstName = encrypt(data.firstName)
    }
    if (data.lastName) {
      updateData.lastName = encrypt(data.lastName)
    }

    const pupil = await prisma.pupil.update({
      where: { admissionNumber },
      data: updateData
    })

    return {
      ...pupil,
      firstName: data.firstName ? data.firstName : decrypt(pupil.firstName),
      lastName: data.lastName ? data.lastName : decrypt(pupil.lastName)
    }
  }

  /**
   * Bulk create pupils
   * Automatically encrypts PII fields
   */
  async bulkCreate(
    pupils: Array<{
      admissionNumber: string
      firstName: string
      lastName: string
      gender: string
      isActive?: boolean
    }>
  ): Promise<number> {
    const encryptedPupils = pupils.map(pupil => ({
      ...pupil,
      firstName: encrypt(pupil.firstName),
      lastName: encrypt(pupil.lastName)
    }))

    const result = await prisma.pupil.createMany({
      data: encryptedPupils,
      skipDuplicates: true
    })

    return result.count
  }

  /**
   * Find pupils by form name
   * Automatically decrypts PII fields
   */
  async findByForm(formName: string): Promise<Pupil[]> {
    const pupils = await prisma.pupil.findMany({
      where: {
        form: formName,
        isActive: true
      },
      orderBy: [{ lastName: 'asc' }, { firstName: 'asc' }]
    })

    return pupils.map(pupil => ({
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }))
  }
}


export const pupilRepository = new PupilRepository()
