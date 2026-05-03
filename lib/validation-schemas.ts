import { z } from 'zod'

/**
 * Validation schemas for server action inputs
 * Using Zod for runtime type validation and error messages
 */

// User Management Schemas
export const CreateUserSchema = z.object({
  username: z.string().min(3, 'Username must be at least 3 characters').max(50),
  password: z.string().min(6, 'Password must be at least 6 characters'),
  roleNames: z.array(z.string()).optional().default([])
})

export const UpdateUserRolesSchema = z.object({
  userId: z.string().min(1, 'User ID is required'),
  roleNames: z.array(z.string()).min(1, 'At least one role is required')
})

export const UpdateUserActiveStatusSchema = z.object({
  userId: z.string().min(1, 'User ID is required'),
  isActive: z.boolean()
})

// Pupil Management Schemas
export const UpdatePupilSchema = z.object({
  admissionNumber: z.string().min(1, 'Admission number is required'),
  data: z.object({
    firstName: z.string().min(1).max(100).optional(),
    lastName: z.string().min(1).max(100).optional(),
    gender: z.enum(['M', 'F']).optional(),
    isActive: z.boolean().optional()
  })
})

export const CreatePupilSchema = z.object({
  admissionNumber: z.string().min(1, 'Admission number is required'),
  firstName: z.string().min(1, 'First name is required').max(100),
  lastName: z.string().min(1, 'Last name is required').max(100),
  gender: z.enum(['M', 'F'], { message: 'Gender must be M or F' }),
  form: z.string().max(20).optional().nullable(),
  isActive: z.boolean().optional().default(true)
})


// Subject Management Schemas
export const CreateSubjectSchema = z.object({
  code: z.string()
    .min(2, 'Subject code must be at least 2 characters')
    .max(10, 'Subject code must be at most 10 characters')
    .regex(/^[A-Z0-9-]+$/, 'Subject code must contain only uppercase letters, numbers, and hyphens'),
  title: z.string().min(1, 'Subject title is required').max(100),
  introduction: z.string().max(500).optional(),
  heOrShe: z.string().max(50).optional(),
  himOrHer: z.string().max(50).optional(),
  hisOrHer: z.string().max(50).optional()
})

export const UpdateSubjectSchema = z.object({
  subjectId: z.string().min(1, 'Subject ID is required'),
  code: z.string()
    .min(2, 'Subject code must be at least 2 characters')
    .max(10, 'Subject code must be at most 10 characters')
    .regex(/^[A-Z0-9-]+$/, 'Subject code must contain only uppercase letters, numbers, and hyphens')
    .optional(),
  title: z.string().min(1).max(100).optional(),
  introduction: z.string().max(500).optional(),
  heOrShe: z.string().max(50).optional(),
  himOrHer: z.string().max(50).optional(),
  hisOrHer: z.string().max(50).optional()
})

export const DeleteSubjectSchema = z.object({
  subjectId: z.string().min(1, 'Subject ID is required')
})

export const AssignUserToSubjectSchema = z.object({
  subjectId: z.string().min(1, 'Subject ID is required'),
  userId: z.string().min(1, 'User ID is required')
})

// Class Management Schemas
export const AssignTeachersSchema = z.object({
  classId: z.string().min(1, 'Class ID is required'),
  teacherIds: z.array(z.string()).min(1, 'At least one teacher is required')
})

// Comment Group Schemas
export const CreateCommentGroupSchema = z.object({
  subjectId: z.string().min(1, 'Subject ID is required'),
  name: z.string().min(1, 'Group name is required').max(100),
  title: z.string().min(1, 'Group title is required').max(200),
  isLinked: z.boolean().optional().default(false),
  linkedField: z.string().max(100).nullable().optional()
}).refine(data => !data.isLinked || (data.linkedField && data.linkedField.trim().length > 0), {
  message: 'Linked field is required when group is linked',
  path: ['linkedField']
})

export const UpdateCommentGroupSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required'),
  name: z.string().min(1).max(100).optional(),
  title: z.string().min(1).max(200).optional(),
  displayOrder: z.number().int().min(0).optional(),
  isLinked: z.boolean().optional(),
  linkedField: z.string().max(100).nullable().optional()
}).refine(data => data.isLinked === undefined || !data.isLinked || (data.linkedField && data.linkedField.trim().length > 0), {
  message: 'Linked field is required when group is linked',
  path: ['linkedField']
})

export const DeleteCommentGroupSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required')
})

// Comment Option Schemas
export const CreateCommentOptionSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required'),
  code: z.string().min(1, 'Code is required').max(20),
  text: z.string().min(1, 'Comment text is required').max(1000)
})

export const UpdateCommentOptionSchema = z.object({
  optionId: z.string().min(1, 'Option ID is required'),
  code: z.string().min(1).max(20).optional(),
  text: z.string().min(1).max(1000).optional()
})

export const DeleteCommentOptionSchema = z.object({
  optionId: z.string().min(1, 'Option ID is required')
})

// Reorder Schemas
export const ReorderCommentGroupsSchema = z.object({
  subjectId: z.string().min(1, 'Subject ID is required'),
  groupIds: z.array(z.string()).min(1, 'At least one group is required')
})

export const ReorderCommentOptionsSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required'),
  optionIds: z.array(z.string()).min(1, 'At least one option is required')
})

// Deadline Management Schemas
export const CreateDeadlineSchema = z.object({
  title: z.string().min(1, 'Title is required').max(200),
  date: z.string().min(1, 'Date is required'),
  description: z.string().max(500).optional()
})

export const UpdateDeadlineSchema = z.object({
  deadlineId: z.string().min(1, 'Deadline ID is required'),
  title: z.string().min(1).max(200).optional(),
  date: z.string().optional(),
  description: z.string().max(500).optional(),
  isActive: z.boolean().optional()
})

export const DeleteDeadlineSchema = z.object({
  deadlineId: z.string().min(1, 'Deadline ID is required')
})

// Common Comment Group Schemas
export const CreateCommonCommentGroupSchema = z.object({
  name: z.string().min(1, 'Group name is required').max(100),
  title: z.string().min(1, 'Group title is required').max(200),
  isLinked: z.boolean().optional().default(false),
  linkedField: z.string().max(100).nullable().optional()
}).refine(data => !data.isLinked || (data.linkedField && data.linkedField.trim().length > 0), {
  message: 'Linked field is required when group is linked',
  path: ['linkedField']
})

export const UpdateCommonCommentGroupSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required'),
  name: z.string().min(1).max(100).optional(),
  title: z.string().min(1).max(200).optional(),
  isLinked: z.boolean().optional(),
  linkedField: z.string().max(100).nullable().optional()
}).refine(data => data.isLinked === undefined || !data.isLinked || (data.linkedField && data.linkedField.trim().length > 0), {
  message: 'Linked field is required when group is linked',
  path: ['linkedField']
})

export const DeleteCommonCommentGroupSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required')
})

export const CreateCommonCommentOptionSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required'),
  code: z.string().min(1, 'Code is required').max(20),
  text: z.string().min(1, 'Comment text is required').max(1000)
})

export const UpdateCommonCommentOptionSchema = z.object({
  optionId: z.string().min(1, 'Option ID is required'),
  code: z.string().min(1).max(20).optional(),
  text: z.string().min(1).max(1000).optional()
})

export const DeleteCommonCommentOptionSchema = z.object({
  optionId: z.string().min(1, 'Option ID is required')
})

export const ReorderCommonCommentGroupsSchema = z.object({
  groupIds: z.array(z.string()).min(1, 'At least one group is required')
})

export const ReorderCommonCommentOptionsSchema = z.object({
  groupId: z.string().min(1, 'Group ID is required'),
  optionIds: z.array(z.string()).min(1, 'At least one option is required')
})

export const UpdateCommentFormatTemplateSchema = z.object({
  value: z.string().min(1, 'Template value is required').max(2000)
})

export const UpdateSubjectCommentFormatSchema = z.object({
  subjectId: z.string().min(1, 'Subject ID is required'),
  commentFormat: z.string().max(500).nullable()
})

/**
 * Helper function to validate form data
 */
export function validateFormData<T>(
  schema: z.ZodSchema<T>,
  data: unknown
): { success: true; data: T } | { success: false; error: string; details: any } {
  const result = schema.safeParse(data)
  
  if (result.success) {
    return { success: true, data: result.data }
  }
  
  return {
    success: false,
    error: 'Validation failed',
    details: result.error.flatten()
  }
}
