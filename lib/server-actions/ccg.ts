"use server"

import { revalidatePath } from 'next/cache'
import { prisma } from '@/lib/prisma'
import { withRole } from '@/lib/auth/with-role'
import { createId } from '@paralleldrive/cuid2'
import { handleServerActionError } from '@/lib/errors'
import { logger } from '@/lib/logger'
import { createAuditLog, logDataChange } from '@/lib/audit-log'
import {
  CreateCommonCommentGroupSchema,
  UpdateCommonCommentGroupSchema,
  DeleteCommonCommentGroupSchema,
  CreateCommonCommentOptionSchema,
  UpdateCommonCommentOptionSchema,
  DeleteCommonCommentOptionSchema,
  ReorderCommonCommentGroupsSchema,
  ReorderCommonCommentOptionsSchema,
  UpdateCommentFormatTemplateSchema,
  validateFormData
} from '@/lib/validation-schemas'

// ============================================================================
// Common Comment Group Management Actions
// ============================================================================

export const createCommonCommentGroup = withRole('admin', async (
  name: string,
  title: string,
  isLinked: boolean = false,
  linkedField: string | null = null
) => {
  try {
    const data = { name, title, isLinked, linkedField }

    const validation = validateFormData(CreateCommonCommentGroupSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    const maxOrder = await (prisma as any).commonCommentGroup.aggregate({
      _max: { displayOrder: true }
    })

    const group = await (prisma as any).commonCommentGroup.create({
      data: {
        id: createId(),
        name: validated.name,
        title: validated.title,
        displayOrder: (maxOrder._max.displayOrder ?? -1) + 1,
        isLinked: validated.isLinked,
        linkedField: validated.isLinked ? validated.linkedField : null
      }
    })

    await createAuditLog({
      action: 'create_common_comment_group',
      entityType: 'common_comment_group',
      entityId: group.id,
      details: { after: { name: validated.name, title: validated.title, isLinked: validated.isLinked, linkedField: validated.linkedField } }
    })

    logger.info('Common comment group created', { name: validated.name })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to create common comment group', { error })
    return handleServerActionError(error)
  }
})

export const updateCommonCommentGroup = withRole('admin', async (
  groupId: string,
  name: string,
  title: string,
  isLinked: boolean = false,
  linkedField: string | null = null
) => {
  try {
    const data = { groupId, name, title, isLinked, linkedField }

    const validation = validateFormData(UpdateCommonCommentGroupSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data
    const currentGroup = await (prisma as any).commonCommentGroup.findUnique({ where: { id: groupId } })

    await (prisma as any).commonCommentGroup.update({
      where: { id: groupId },
      data: {
        name: validated.name,
        title: validated.title,
        isLinked: validated.isLinked ?? false,
        linkedField: validated.isLinked ? validated.linkedField : null,
      }
    })

    await logDataChange(
      'update_common_comment_group',
      'common_comment_group',
      groupId,
      currentGroup ? { name: currentGroup.name, title: currentGroup.title, isLinked: currentGroup.isLinked, linkedField: currentGroup.linkedField } : null,
      { name: validated.name, title: validated.title, isLinked: validated.isLinked, linkedField: validated.linkedField }
    )

    logger.info('Common comment group updated', { groupId })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to update common comment group', { error, groupId })
    return handleServerActionError(error)
  }
})

export const deleteCommonCommentGroup = withRole('admin', async (
  groupId: string
) => {
  try {
    const validation = validateFormData(DeleteCommonCommentGroupSchema, { groupId })
    if (!validation.success) {
      return validation
    }

    const currentGroup = await (prisma as any).commonCommentGroup.findUnique({ where: { id: groupId } })

    await (prisma as any).commonCommentGroup.delete({
      where: { id: groupId }
    })

    await createAuditLog({
      action: 'delete_common_comment_group',
      entityType: 'common_comment_group',
      entityId: groupId,
      details: currentGroup ? { before: { name: currentGroup.name, title: currentGroup.title } } : undefined
    })

    logger.info('Common comment group deleted', { groupId })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete common comment group', { error, groupId })
    return handleServerActionError(error)
  }
})

export const reorderCommonCommentGroups = withRole('admin', async (
  items: { id: string; order: number }[]
) => {
  try {
    const validation = validateFormData(ReorderCommonCommentGroupsSchema, {
      groupIds: items.map(i => i.id)
    })
    if (!validation.success) {
      return validation
    }

    await Promise.all(
      items.map(item =>
        (prisma as any).commonCommentGroup.update({
          where: { id: item.id },
          data: { displayOrder: item.order }
        })
      )
    )

    await createAuditLog({
      action: 'reorder_common_comment_groups',
      entityType: 'common_comment_group',
      details: { newOrder: items }
    })

    logger.info('Common comment groups reordered', { count: items.length })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to reorder common comment groups', { error })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Common Comment Option Management Actions
// ============================================================================

export const createCommonCommentOption = withRole('admin', async (
  groupId: string,
  code: string,
  text: string
) => {
  try {
    const data = { groupId, code, text }

    const validation = validateFormData(CreateCommonCommentOptionSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    const maxOrder = await (prisma as any).commonCommentOption.aggregate({
      where: { groupId },
      _max: { displayOrder: true }
    })

    const option = await (prisma as any).commonCommentOption.create({
      data: {
        id: createId(),
        code: validated.code,
        text: validated.text,
        groupId: validated.groupId,
        displayOrder: (maxOrder._max.displayOrder ?? -1) + 1
      }
    })

    await createAuditLog({
      action: 'create_common_comment_option',
      entityType: 'common_comment_option',
      entityId: option.id,
      details: { after: { code: validated.code, text: validated.text, groupId } }
    })

    logger.info('Common comment option created', { groupId, code: validated.code })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to create common comment option', { error, groupId })
    return handleServerActionError(error)
  }
})

export const updateCommonCommentOption = withRole('admin', async (
  optionId: string,
  code: string,
  text: string
) => {
  try {
    const data = { optionId, code, text }

    const validation = validateFormData(UpdateCommonCommentOptionSchema, data)
    if (!validation.success) {
      return validation
    }

    const validated = validation.data
    const currentOption = await (prisma as any).commonCommentOption.findUnique({ where: { id: optionId } })

    await (prisma as any).commonCommentOption.update({
      where: { id: optionId },
      data: {
        code: validated.code,
        text: validated.text,
      }
    })

    await logDataChange(
      'update_common_comment_option',
      'common_comment_option',
      optionId,
      currentOption ? { code: currentOption.code, text: currentOption.text } : null,
      { code: validated.code, text: validated.text }
    )

    logger.info('Common comment option updated', { optionId })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to update common comment option', { error, optionId })
    return handleServerActionError(error)
  }
})

export const deleteCommonCommentOption = withRole('admin', async (
  optionId: string
) => {
  try {
    const validation = validateFormData(DeleteCommonCommentOptionSchema, { optionId })
    if (!validation.success) {
      return validation
    }

    const currentOption = await (prisma as any).commonCommentOption.findUnique({ where: { id: optionId } })

    await (prisma as any).commonCommentOption.delete({
      where: { id: optionId }
    })

    await createAuditLog({
      action: 'delete_common_comment_option',
      entityType: 'common_comment_option',
      entityId: optionId,
      details: currentOption ? { before: { code: currentOption.code, text: currentOption.text } } : undefined
    })

    logger.info('Common comment option deleted', { optionId })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to delete common comment option', { error, optionId })
    return handleServerActionError(error)
  }
})

export const reorderCommonCommentOptions = withRole('admin', async (
  groupId: string,
  items: { id: string; order: number }[]
) => {
  try {
    const validation = validateFormData(ReorderCommonCommentOptionsSchema, {
      groupId,
      optionIds: items.map(i => i.id)
    })
    if (!validation.success) {
      return validation
    }

    await Promise.all(
      items.map(item =>
        (prisma as any).commonCommentOption.update({
          where: { id: item.id },
          data: { displayOrder: item.order }
        })
      )
    )

    await createAuditLog({
      action: 'reorder_common_comment_options',
      entityType: 'common_comment_option',
      details: { groupId, newOrder: items }
    })

    logger.info('Common comment options reordered', { groupId, count: items.length })
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to reorder common comment options', { error, groupId })
    return handleServerActionError(error)
  }
})

// ============================================================================
// Comment Format Template Management
// ============================================================================

export const updateCommentFormatTemplate = withRole('admin', async (
  value: string
) => {
  try {
    const validation = validateFormData(UpdateCommentFormatTemplateSchema, { value })
    if (!validation.success) {
      return validation
    }

    const validated = validation.data

    await (prisma as any).appSetting.upsert({
      where: { key: 'comment_format_template' },
      update: { value: validated.value },
      create: { key: 'comment_format_template', value: validated.value }
    })

    await createAuditLog({
      action: 'update_comment_format_template',
      entityType: 'app_setting',
      entityId: 'comment_format_template',
      details: { value: validated.value }
    })

    logger.info('Comment format template updated')
    revalidatePath('/admin')

    return { success: true }
  } catch (error) {
    logger.error('Failed to update comment format template', { error })
    return handleServerActionError(error)
  }
})
