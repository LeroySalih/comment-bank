"use server"

import { revalidatePath } from 'next/cache'
import { pool } from '@/lib/db'
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

    const { rows: maxRows } = await pool.query<{ max: number | null }>(
      `SELECT MAX("displayOrder") as max FROM "CommonCommentGroup"`
    )
    const nextOrder = (maxRows[0].max ?? -1) + 1

    const { rows: groupRows } = await pool.query(
      `INSERT INTO "CommonCommentGroup" (id, name, title, "displayOrder", "isLinked", "linkedField")
       VALUES ($1, $2, $3, $4, $5, $6) RETURNING *`,
      [
        createId(),
        validated.name,
        validated.title,
        nextOrder,
        validated.isLinked,
        validated.isLinked ? validated.linkedField : null
      ]
    )
    const group = groupRows[0]

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

    const { rows: currentRows } = await pool.query(
      `SELECT * FROM "CommonCommentGroup" WHERE id = $1`,
      [groupId]
    )
    const currentGroup = currentRows[0] ?? null

    await pool.query(
      `UPDATE "CommonCommentGroup"
       SET name = $1, title = $2, "isLinked" = $3, "linkedField" = $4
       WHERE id = $5`,
      [
        validated.name,
        validated.title,
        validated.isLinked ?? false,
        validated.isLinked ? validated.linkedField : null,
        groupId
      ]
    )

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

    const { rows: currentRows } = await pool.query(
      `SELECT * FROM "CommonCommentGroup" WHERE id = $1`,
      [groupId]
    )
    const currentGroup = currentRows[0] ?? null

    await pool.query(`DELETE FROM "CommonCommentGroup" WHERE id = $1`, [groupId])

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

    const client = await pool.connect()
    try {
      await client.query('BEGIN')
      for (const item of items) {
        await client.query(
          `UPDATE "CommonCommentGroup" SET "displayOrder" = $1 WHERE id = $2`,
          [item.order, item.id]
        )
      }
      await client.query('COMMIT')
    } catch (err) {
      await client.query('ROLLBACK')
      throw err
    } finally {
      client.release()
    }

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

    const { rows: maxRows } = await pool.query<{ max: number | null }>(
      `SELECT MAX("displayOrder") as max FROM "CommonCommentOption" WHERE "groupId" = $1`,
      [groupId]
    )
    const nextOrder = (maxRows[0].max ?? -1) + 1

    const { rows: optionRows } = await pool.query(
      `INSERT INTO "CommonCommentOption" (id, code, text, "groupId", "displayOrder")
       VALUES ($1, $2, $3, $4, $5) RETURNING *`,
      [
        createId(),
        validated.code,
        validated.text,
        validated.groupId,
        nextOrder
      ]
    )
    const option = optionRows[0]

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

    const { rows: currentRows } = await pool.query(
      `SELECT * FROM "CommonCommentOption" WHERE id = $1`,
      [optionId]
    )
    const currentOption = currentRows[0] ?? null

    await pool.query(
      `UPDATE "CommonCommentOption" SET code = $1, text = $2 WHERE id = $3`,
      [validated.code, validated.text, optionId]
    )

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

    const { rows: currentRows } = await pool.query(
      `SELECT * FROM "CommonCommentOption" WHERE id = $1`,
      [optionId]
    )
    const currentOption = currentRows[0] ?? null

    await pool.query(`DELETE FROM "CommonCommentOption" WHERE id = $1`, [optionId])

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

    const client = await pool.connect()
    try {
      await client.query('BEGIN')
      for (const item of items) {
        await client.query(
          `UPDATE "CommonCommentOption" SET "displayOrder" = $1 WHERE id = $2`,
          [item.order, item.id]
        )
      }
      await client.query('COMMIT')
    } catch (err) {
      await client.query('ROLLBACK')
      throw err
    } finally {
      client.release()
    }

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

    await pool.query(
      `INSERT INTO "AppSetting" (key, value) VALUES ($1, $2)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      ['comment_format_template', validated.value]
    )

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
