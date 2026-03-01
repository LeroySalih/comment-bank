import { pool } from "@/lib/db"
import { AdminTabs } from "./_components/admin-tabs"
import SignOutButton from "@/components/SignOutButton"

export const dynamic = 'force-dynamic'

export default async function AdminPage() {
  // Fetch users with their roles
  const { rows: userRows } = await pool.query(
    `SELECT u.id, u.username, u."isActive",
            coalesce(json_agg(json_build_object('id', r.id, 'name', r.name)) FILTER (WHERE r.id IS NOT NULL), '[]') as "Role"
     FROM "User" u
     LEFT JOIN "_RoleToUser" ru ON ru."B" = u.id
     LEFT JOIN "Role" r ON r.id = ru."A"
     GROUP BY u.id, u.username, u."isActive"
     ORDER BY u.username ASC`
  )
  const users = userRows

  const { rows: roles } = await pool.query(
    `SELECT * FROM "Role" ORDER BY name ASC`
  )

  // Fetch subjects with counts and comment groups
  const { rows: subjectRows } = await pool.query(
    `SELECT s.*,
            (SELECT COUNT(*) FROM "Class" c WHERE c."subjectId" = s.id) as class_count,
            (SELECT COUNT(*) FROM "CommentGroup" cg WHERE cg."subjectId" = s.id) as group_count
     FROM "Subject" s
     ORDER BY s.code ASC`
  )

  // Fetch users assigned to each subject
  const subjectIds = subjectRows.map((s: any) => s.id)
  let subjectUserMap = new Map<string, any[]>()
  if (subjectIds.length > 0) {
    const { rows: subjectUserRows } = await pool.query(
      `SELECT su."A" as subject_id, u.id, u.username,
              coalesce(json_agg(json_build_object('name', r.name)) FILTER (WHERE r.id IS NOT NULL), '[]') as roles
       FROM "_SubjectToUser" su
       JOIN "User" u ON u.id = su."B"
       LEFT JOIN "_RoleToUser" ru ON ru."B" = u.id
       LEFT JOIN "Role" r ON r.id = ru."A"
       WHERE su."A" = ANY($1::text[])
       GROUP BY su."A", u.id, u.username`,
      [subjectIds]
    )
    for (const row of subjectUserRows) {
      const arr = subjectUserMap.get(row.subject_id) ?? []
      arr.push({ id: row.id, username: row.username, Role: row.roles })
      subjectUserMap.set(row.subject_id, arr)
    }
  }

  // Fetch comment groups with options per subject
  let subjectGroupMap = new Map<string, any[]>()
  if (subjectIds.length > 0) {
    const { rows: groupRows } = await pool.query(
      `SELECT cg.* FROM "CommentGroup" cg WHERE cg."subjectId" = ANY($1::text[]) ORDER BY cg."displayOrder" ASC`,
      [subjectIds]
    )
    const groupIds = groupRows.map((g: any) => g.id)
    let optionsByGroup = new Map<string, any[]>()
    if (groupIds.length > 0) {
      const { rows: optRows } = await pool.query(
        `SELECT * FROM "CommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
        [groupIds]
      )
      for (const opt of optRows) {
        const arr = optionsByGroup.get(opt.groupId) ?? []
        arr.push(opt)
        optionsByGroup.set(opt.groupId, arr)
      }
    }
    for (const g of groupRows) {
      (g as any).CommentOption = optionsByGroup.get(g.id) ?? []
      const arr = subjectGroupMap.get(g.subjectId) ?? []
      arr.push(g)
      subjectGroupMap.set(g.subjectId, arr)
    }
  }

  const subjects = subjectRows.map((s: any) => ({
    ...s,
    _count: {
      Class: Number(s.class_count),
      CommentGroup: Number(s.group_count)
    },
    User: subjectUserMap.get(s.id) ?? [],
    CommentGroup: subjectGroupMap.get(s.id) ?? []
  }))

  const { rows: deadlines } = await pool.query(
    `SELECT * FROM "Deadline" ORDER BY date ASC`
  )

  const { rows: classes } = await pool.query(
    `SELECT id, name FROM "Class" ORDER BY name ASC`
  )

  // Fetch common comment groups with options
  const { rows: cgRows } = await pool.query(
    `SELECT * FROM "CommonCommentGroup" ORDER BY "displayOrder" ASC`
  )
  if (cgRows.length > 0) {
    const cgIds = cgRows.map((g: any) => g.id)
    const { rows: cgOptRows } = await pool.query(
      `SELECT * FROM "CommonCommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
      [cgIds]
    )
    const cgOptsByGroup = new Map<string, any[]>()
    for (const opt of cgOptRows) {
      const arr = cgOptsByGroup.get(opt.groupId) ?? []
      arr.push(opt)
      cgOptsByGroup.set(opt.groupId, arr)
    }
    for (const g of cgRows) {
      (g as any).CommonCommentOption = cgOptsByGroup.get(g.id) ?? []
    }
  }
  const commonGroups = cgRows

  const { rows: settingRows } = await pool.query(
    `SELECT value FROM "AppSetting" WHERE key = 'comment_format_template'`
  )
  const commentFormatTemplate = settingRows[0]?.value || ''

  return (
    <div className="container mx-auto py-10">
      <div className="flex justify-between items-center mb-8">
        <h1 className="text-3xl font-bold">Admin Dashboard</h1>
        <SignOutButton />
      </div>

      <AdminTabs users={users} roles={roles} subjects={subjects} deadlines={deadlines} classes={classes} commonGroups={commonGroups} commentFormatTemplate={commentFormatTemplate} />
    </div>
  )
}
