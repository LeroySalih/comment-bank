import { describe, it, expect } from 'vitest'
import {
  CreateSubjectSchema,
  UpdatePupilSchema,
  CreateCommentGroupSchema,
  validateFormData
} from '../validation-schemas'

describe('Validation Schemas', () => {
  describe('CreateSubjectSchema', () => {
    it('should validate correct subject data', () => {
      const validData = {
        code: 'CS101',
        title: 'Computer Science',
        introduction: 'Introduction to CS'
      }
      
      const result = CreateSubjectSchema.safeParse(validData)
      expect(result.success).toBe(true)
    })

    it('should reject invalid subject code', () => {
      const invalidData = {
        code: 'cs101', // lowercase not allowed
        title: 'Computer Science'
      }
      
      const result = CreateSubjectSchema.safeParse(invalidData)
      expect(result.success).toBe(false)
    })

    it('should reject subject code with special characters', () => {
      const invalidData = {
        code: 'CS-101', // hyphen not allowed
        title: 'Computer Science'
      }
      
      const result = CreateSubjectSchema.safeParse(invalidData)
      expect(result.success).toBe(false)
    })

    it('should reject too short subject code', () => {
      const invalidData = {
        code: 'C',
        title: 'Computer Science'
      }
      
      const result = CreateSubjectSchema.safeParse(invalidData)
      expect(result.success).toBe(false)
    })

    it('should reject missing title', () => {
      const invalidData = {
        code: 'CS101'
      }
      
      const result = CreateSubjectSchema.safeParse(invalidData)
      expect(result.success).toBe(false)
    })
  })

  describe('UpdatePupilSchema', () => {
    it('should validate correct pupil update', () => {
      const validData = {
        admissionNumber: '12345',
        data: {
          firstName: 'John',
          lastName: 'Doe',
          gender: 'M',
          isActive: true
        }
      }
      
      const result = UpdatePupilSchema.safeParse(validData)
      expect(result.success).toBe(true)
    })

    it('should allow partial updates', () => {
      const validData = {
        admissionNumber: '12345',
        data: {
          firstName: 'John'
        }
      }
      
      const result = UpdatePupilSchema.safeParse(validData)
      expect(result.success).toBe(true)
    })

    it('should reject invalid gender', () => {
      const invalidData = {
        admissionNumber: '12345',
        data: {
          gender: 'X' // Only M or F allowed
        }
      }
      
      const result = UpdatePupilSchema.safeParse(invalidData)
      expect(result.success).toBe(false)
    })
  })

  describe('CreateCommentGroupSchema', () => {
    it('should validate correct comment group', () => {
      const validData = {
        subjectId: 'subject-123',
        name: 'WP',
        title: 'Written Production'
      }
      
      const result = CreateCommentGroupSchema.safeParse(validData)
      expect(result.success).toBe(true)
    })

    it('should reject missing required fields', () => {
      const invalidData = {
        subjectId: 'subject-123',
        name: 'WP'
        // title missing
      }
      
      const result = CreateCommentGroupSchema.safeParse(invalidData)
      expect(result.success).toBe(false)
    })

    it('should reject empty strings', () => {
      const invalidData = {
        subjectId: 'subject-123',
        name: '',
        title: 'Written Production'
      }
      
      const result = CreateCommentGroupSchema.safeParse(invalidData)
      expect(result.success).toBe(false)
    })
  })

  describe('validateFormData helper', () => {
    it('should return success for valid data', () => {
      const data = {
        code: 'CS101',
        title: 'Computer Science'
      }
      
      const result = validateFormData(CreateSubjectSchema, data)
      expect(result.success).toBe(true)
      if (result.success) {
        expect(result.data).toEqual(expect.objectContaining(data))
      }
    })

    it('should return error details for invalid data', () => {
      const data = {
        code: 'cs', // too short and lowercase
        title: ''  // empty
      }
      
      const result = validateFormData(CreateSubjectSchema, data)
      expect(result.success).toBe(false)
      if (!result.success) {
        expect(result.error).toBe('Validation failed')
        expect(result.details).toBeDefined()
      }
    })
  })
})
