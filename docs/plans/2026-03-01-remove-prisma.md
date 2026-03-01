# Remove Prisma — Use pg Directly

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Replace Prisma's ORM client with direct `pg` Pool queries across all 14 files, keeping repository interfaces and server-action signatures identical.

**Architecture:** The `pg` package is already installed. We create a `Pool` singleton in `lib/db.ts`, define TS interfaces in `lib/types/db.ts` to replace Prisma-generated types, then rewrite each file file-by-file, replacing `prisma.*` with `pool.query()` SQL. The repository class boundaries and all method signatures are preserved.

**Tech Stack:** `pg` (already installed), TypeScript, Next.js Server Actions, PostgreSQL

---

## Pre-flight: Confirm junction table names

Before writing any queries, confirm Prisma's implicit many-to-many junction table names in the database. Run this once and record the output — you'll need these table names in Tasks 4, 5, 6.

```bash
psql $DATABASE_URL -c "\dt" | grep -E "^(public|) *\| *_"
```

Expected tables:
- `_ClassToUser` — User ↔ Class (many-to-many)
- `_RoleToUser` — User ↔ Role (many-to-many)
- `_SubjectToUser` — User ↔ Subject (many-to-many)

Also confirm column names (`A`, `B`) by running:
```bash
psql $DATABASE_URL -c "\d \"_RoleToUser\""
```

Prisma always uses columns `A` (the alphabetically-first model ID) and `B` (the other). For `_RoleToUser`: `A` = `roleId`, `B` = `userId`.

---

## Task 1: Create pg pool singleton (`lib/db.ts`)

**Files:**
- Create: `lib/db.ts`
- Delete: `lib/prisma.ts` (at end of task)

**Step 1: Create the new pool singleton**

```typescript
// lib/db.ts
import { Pool } from 'pg'

function createPool() {
  return new Pool({ connectionString: process.env.DATABASE_URL! })
}

const globalForPool = globalThis as unknown as { pool: Pool }

export const pool = globalForPool.pool || createPool()

if (process.env.NODE_ENV !== 'production') globalForPool.pool = pool
```

**Step 2: Verify it compiles**

```bash
npx tsc --noEmit --incremental false 2>&1 | head -20
```

Expected: Only errors about files still importing from `@/lib/prisma` — that's fine, we'll fix those in subsequent tasks.

**Step 3: Commit**

```bash
git add lib/db.ts
git commit -m "feat: add pg pool singleton to replace PrismaClient"
```

---

## Task 2: Create shared TypeScript interfaces (`lib/types/db.ts`)

Prisma generated types like `User`, `Class`, `Assignment` etc. from the schema. With raw `pg`, we define these ourselves.

**Files:**
- Create: `lib/types/db.ts`

**Step 1: Write the interfaces**

```typescript
// lib/types/db.ts

export interface DbUser {
  id: string
  username: string
  password: string
  isActive: boolean
}

export interface DbRole {
  id: string
  name: string
}

export interface DbUserWithRoles extends DbUser {
  Role: DbRole[]
}

export interface DbSubject {
  id: string
  code: string
  title: string | null
  studiedComment: string | null
  commentFormat: string | null
}

export interface DbClass {
  id: string
  name: string
  year: string | null
  subjectId: string
}

export interface DbPupil {
  admissionNumber: string
  firstName: string
  lastName: string
  gender: string
  isActive: boolean
  form: string | null
}

export interface DbAssignment {
  id: string
  pupilId: string
  classId: string
  eoyLevel: string | null
  targetLevel: string | null
  actualLevel: string | null
  finalComment: string | null
  linkedData: Record<string, any> | null
  checkStatus: string
  checkNote: string | null
  checkedAt: Date | null
  checkedById: string | null
}

export interface DbCommentGroup {
  id: string
  name: string
  displayOrder: number
  subjectId: string
  title: string
  isLinked: boolean
  linkedField: string | null
}

export interface DbCommentOption {
  id: string
  code: string
  text: string
  displayOrder: number
  groupId: string
}

export interface DbPupilCode {
  id: string
  assignmentId: string
  groupId: string
  code: string | null
}

export interface DbCommonCommentGroup {
  id: string
  name: string
  title: string
  displayOrder: number
  isLinked: boolean
  linkedField: string | null
}

export interface DbCommonCommentOption {
  id: string
  code: string
  text: string
  displayOrder: number
  groupId: string
}

export interface DbCommonPupilCode {
  id: string
  assignmentId: string
  commonGroupId: string
  code: string | null
}

export interface DbDeadline {
  id: string
  title: string
  date: Date
  description: string | null
  isActive: boolean
  createdAt: Date
}

export interface DbAppSetting {
  key: string
  value: string
}

export interface DbAuditLog {
  id: string
  userId: string | null
  username: string | null
  action: string
  entityType: string | null
  entityId: string | null
  details: string | null
  ipAddress: string | null
  userAgent: string | null
  createdAt: Date
}
```

**Step 2: Commit**

```bash
git add lib/types/db.ts
git commit -m "feat: add shared db type interfaces to replace Prisma-generated types"
```

---

## Task 3: Rewrite `lib/db/repositories/pupil-repository.ts`

The `Pupil` table maps directly. `findAll` with a query searches encrypted values — keep that behaviour as-is (it's pre-existing). The `import type { Pupil } from '@prisma/client'` gets replaced with `DbPupil` from our new types file.

**Files:**
- Modify: `lib/db/repositories/pupil-repository.ts`

**Step 1: Rewrite the file**

```typescript
import { pool } from '@/lib/db'
import { encrypt, decrypt } from '@/lib/encryption'
import { NotFoundError } from '@/lib/errors'
import type { DbPupil } from '@/lib/types/db'

export class PupilRepository {
  async findAll(query?: string): Promise<DbPupil[]> {
    let sql: string
    let params: any[]

    if (query) {
      sql = `
        SELECT * FROM "Pupil"
        WHERE "firstName" ILIKE $1 OR "lastName" ILIKE $1 OR "admissionNumber" ILIKE $1
        ORDER BY "lastName" ASC, "firstName" ASC
      `
      params = [`%${query}%`]
    } else {
      sql = `SELECT * FROM "Pupil" ORDER BY "lastName" ASC, "firstName" ASC`
      params = []
    }

    const { rows } = await pool.query<DbPupil>(sql, params)

    return rows.map(pupil => ({
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }))
  }

  async findByAdmissionNumber(admissionNumber: string): Promise<DbPupil | null> {
    const { rows } = await pool.query<DbPupil>(
      `SELECT * FROM "Pupil" WHERE "admissionNumber" = $1`,
      [admissionNumber]
    )

    if (rows.length === 0) return null

    const pupil = rows[0]
    return {
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }
  }

  async create(data: {
    admissionNumber: string
    firstName: string
    lastName: string
    gender: string
    isActive?: boolean
  }): Promise<DbPupil> {
    const { rows } = await pool.query<DbPupil>(
      `INSERT INTO "Pupil" ("admissionNumber", "firstName", "lastName", "gender", "isActive")
       VALUES ($1, $2, $3, $4, $5)
       RETURNING *`,
      [
        data.admissionNumber,
        encrypt(data.firstName),
        encrypt(data.lastName),
        data.gender,
        data.isActive ?? true
      ]
    )

    return {
      ...rows[0],
      firstName: data.firstName,
      lastName: data.lastName
    }
  }

  async update(
    admissionNumber: string,
    data: {
      firstName?: string
      lastName?: string
      gender?: string
      isActive?: boolean
      form?: string | null
    }
  ): Promise<DbPupil> {
    const sets: string[] = []
    const params: any[] = []
    let idx = 1

    if (data.firstName !== undefined) {
      sets.push(`"firstName" = $${idx++}`)
      params.push(encrypt(data.firstName))
    }
    if (data.lastName !== undefined) {
      sets.push(`"lastName" = $${idx++}`)
      params.push(encrypt(data.lastName))
    }
    if (data.gender !== undefined) {
      sets.push(`"gender" = $${idx++}`)
      params.push(data.gender)
    }
    if (data.isActive !== undefined) {
      sets.push(`"isActive" = $${idx++}`)
      params.push(data.isActive)
    }
    if (data.form !== undefined) {
      sets.push(`"form" = $${idx++}`)
      params.push(data.form)
    }

    params.push(admissionNumber)

    const { rows } = await pool.query<DbPupil>(
      `UPDATE "Pupil" SET ${sets.join(', ')} WHERE "admissionNumber" = $${idx} RETURNING *`,
      params
    )

    return {
      ...rows[0],
      firstName: data.firstName ?? decrypt(rows[0].firstName),
      lastName: data.lastName ?? decrypt(rows[0].lastName)
    }
  }

  async bulkCreate(
    pupils: Array<{
      admissionNumber: string
      firstName: string
      lastName: string
      gender: string
      form?: string | null
      isActive?: boolean
    }>
  ): Promise<number> {
    if (pupils.length === 0) return 0

    const values = pupils.map((p, i) => {
      const base = i * 6
      return `($${base + 1}, $${base + 2}, $${base + 3}, $${base + 4}, $${base + 5}, $${base + 6})`
    }).join(', ')

    const params = pupils.flatMap(p => [
      p.admissionNumber,
      encrypt(p.firstName),
      encrypt(p.lastName),
      p.gender,
      p.form ?? null,
      p.isActive ?? true
    ])

    const { rowCount } = await pool.query(
      `INSERT INTO "Pupil" ("admissionNumber", "firstName", "lastName", "gender", "form", "isActive")
       VALUES ${values}
       ON CONFLICT ("admissionNumber") DO NOTHING`,
      params
    )

    return rowCount ?? 0
  }

  async findByForm(formName: string): Promise<DbPupil[]> {
    const { rows } = await pool.query<DbPupil>(
      `SELECT * FROM "Pupil" WHERE "form" = $1 AND "isActive" = true ORDER BY "lastName" ASC, "firstName" ASC`,
      [formName]
    )

    return rows.map(pupil => ({
      ...pupil,
      firstName: decrypt(pupil.firstName),
      lastName: decrypt(pupil.lastName)
    }))
  }
}

export const pupilRepository = new PupilRepository()
```

**Step 2: Verify TypeScript**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "pupil-repository"
```

Expected: No errors for this file.

**Step 3: Commit**

```bash
git add lib/db/repositories/pupil-repository.ts
git commit -m "feat: replace Prisma with pg in pupil-repository"
```

---

## Task 4: Rewrite `lib/db/repositories/user-repository.ts`

This file manages User + Role with a many-to-many. The junction table is `_RoleToUser` with columns `A` (roleId) and `B` (userId). We do a DELETE + INSERT for `updateRoles` (simpler than upsert).

**Files:**
- Modify: `lib/db/repositories/user-repository.ts`

**Step 1: Rewrite the file**

```typescript
import { pool } from '@/lib/db'
import { NotFoundError } from '@/lib/errors'
import { createId } from '@paralleldrive/cuid2'
import type { DbUser, DbRole, DbUserWithRoles } from '@/lib/types/db'

// Helper: attach roles to a user row
async function attachRoles(userId: string): Promise<DbRole[]> {
  const { rows } = await pool.query<DbRole>(
    `SELECT r.* FROM "Role" r
     JOIN "_RoleToUser" ru ON ru."A" = r.id
     WHERE ru."B" = $1`,
    [userId]
  )
  return rows
}

export class UserRepository {
  async findAll(): Promise<DbUserWithRoles[]> {
    const { rows } = await pool.query<DbUser>(
      `SELECT * FROM "User" ORDER BY "username" ASC`
    )

    return Promise.all(
      rows.map(async user => ({
        ...user,
        Role: await attachRoles(user.id)
      }))
    )
  }

  async findById(id: string): Promise<DbUserWithRoles & { Subject: any[]; Class: any[] }> {
    const { rows } = await pool.query<DbUser>(
      `SELECT * FROM "User" WHERE id = $1`,
      [id]
    )

    if (rows.length === 0) {
      throw new NotFoundError(`User with ID ${id} not found`)
    }

    const [roles, subjects, classes] = await Promise.all([
      attachRoles(id),
      pool.query(`SELECT * FROM "Subject" s JOIN "_SubjectToUser" su ON su."A" = s.id WHERE su."B" = $1`, [id]),
      pool.query(`SELECT * FROM "Class" c JOIN "_ClassToUser" cu ON cu."B" = c.id WHERE cu."A" = $1`, [id])
    ])

    return {
      ...rows[0],
      Role: roles,
      Subject: subjects.rows,
      Class: classes.rows
    }
  }

  async findByUsername(username: string): Promise<(DbUserWithRoles) | null> {
    const { rows } = await pool.query<DbUser>(
      `SELECT * FROM "User" WHERE "username" = $1`,
      [username]
    )

    if (rows.length === 0) return null

    return {
      ...rows[0],
      Role: await attachRoles(rows[0].id)
    }
  }

  async create(data: {
    username: string
    password: string
    roleNames?: string[]
  }): Promise<DbUserWithRoles> {
    const { roleNames = [], ...userData } = data
    const userId = createId()

    const client = await pool.connect()
    try {
      await client.query('BEGIN')

      await client.query(
        `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)`,
        [userId, userData.username, userData.password]
      )

      for (const name of roleNames) {
        // Ensure role exists
        await client.query(
          `INSERT INTO "Role" (id, name) VALUES ($1, $2) ON CONFLICT (name) DO NOTHING`,
          [name, name]
        )
        // Link user to role
        const { rows: roleRows } = await client.query<DbRole>(
          `SELECT id FROM "Role" WHERE name = $1`,
          [name]
        )
        await client.query(
          `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [roleRows[0].id, userId]
        )
      }

      await client.query('COMMIT')
    } catch (e) {
      await client.query('ROLLBACK')
      throw e
    } finally {
      client.release()
    }

    const { rows } = await pool.query<DbUser>(`SELECT * FROM "User" WHERE id = $1`, [userId])
    return {
      ...rows[0],
      Role: await attachRoles(userId)
    }
  }

  async updateRoles(userId: string, roleNames: string[]): Promise<DbUserWithRoles> {
    const client = await pool.connect()
    try {
      await client.query('BEGIN')

      // Remove all existing roles for this user
      await client.query(`DELETE FROM "_RoleToUser" WHERE "B" = $1`, [userId])

      for (const name of roleNames) {
        await client.query(
          `INSERT INTO "Role" (id, name) VALUES ($1, $2) ON CONFLICT (name) DO NOTHING`,
          [name, name]
        )
        const { rows } = await client.query<DbRole>(
          `SELECT id FROM "Role" WHERE name = $1`,
          [name]
        )
        await client.query(
          `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [rows[0].id, userId]
        )
      }

      await client.query('COMMIT')
    } catch (e) {
      await client.query('ROLLBACK')
      throw e
    } finally {
      client.release()
    }

    const { rows } = await pool.query<DbUser>(`SELECT * FROM "User" WHERE id = $1`, [userId])
    return {
      ...rows[0],
      Role: await attachRoles(userId)
    }
  }

  async delete(userId: string): Promise<void> {
    await pool.query(`DELETE FROM "User" WHERE id = $1`, [userId])
  }
}

export const userRepository = new UserRepository()
```

**Step 2: Verify TypeScript**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "user-repository"
```

Expected: No errors for this file.

**Step 3: Commit**

```bash
git add lib/db/repositories/user-repository.ts
git commit -m "feat: replace Prisma with pg in user-repository"
```

---

## Task 5: Rewrite `lib/db/repositories/subject-repository.ts`

Subject has User (HOD, many-to-many via `_SubjectToUser`) and Class (one-to-many). `findAll` returns classes + user HODs + a comment group count.

**Files:**
- Modify: `lib/db/repositories/subject-repository.ts`

**Step 1: Rewrite the file**

```typescript
import { pool } from '@/lib/db'
import { NotFoundError } from '@/lib/errors'
import { createId } from '@paralleldrive/cuid2'
import type { DbSubject, DbClass, DbCommentGroup, DbCommentOption } from '@/lib/types/db'

export class SubjectRepository {
  async findAll() {
    const { rows: subjects } = await pool.query<DbSubject>(
      `SELECT * FROM "Subject" ORDER BY code ASC`
    )

    return Promise.all(subjects.map(async subject => {
      const [classes, users, countResult] = await Promise.all([
        pool.query<DbClass>(
          `SELECT * FROM "Class" WHERE "subjectId" = $1 ORDER BY name ASC`,
          [subject.id]
        ),
        pool.query(
          `SELECT u.id, u.username FROM "User" u
           JOIN "_SubjectToUser" su ON su."B" = u.id
           WHERE su."A" = $1`,
          [subject.id]
        ),
        pool.query<{ count: string }>(
          `SELECT COUNT(*) as count FROM "CommentGroup" WHERE "subjectId" = $1`,
          [subject.id]
        )
      ])

      return {
        ...subject,
        Class: classes.rows,
        User: users.rows,
        _count: { CommentGroup: parseInt(countResult.rows[0].count) }
      }
    }))
  }

  async findByIdWithDetails(id: string) {
    const { rows } = await pool.query<DbSubject>(
      `SELECT * FROM "Subject" WHERE id = $1`,
      [id]
    )

    if (rows.length === 0) {
      throw new NotFoundError(`Subject with ID ${id} not found`)
    }

    const subject = rows[0]

    const [classes, users, commentGroups] = await Promise.all([
      pool.query<DbClass>(
        `SELECT * FROM "Class" WHERE "subjectId" = $1 ORDER BY name ASC`,
        [id]
      ),
      pool.query(
        `SELECT u.id, u.username FROM "User" u
         JOIN "_SubjectToUser" su ON su."B" = u.id
         WHERE su."A" = $1`,
        [id]
      ),
      pool.query<DbCommentGroup>(
        `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 ORDER BY "displayOrder" ASC`,
        [id]
      )
    ])

    const groups = await Promise.all(commentGroups.rows.map(async group => {
      const { rows: options } = await pool.query<DbCommentOption>(
        `SELECT * FROM "CommentOption" WHERE "groupId" = $1 ORDER BY "displayOrder" ASC`,
        [group.id]
      )
      return { ...group, CommentOption: options }
    }))

    return {
      ...subject,
      Class: classes.rows,
      User: users.rows,
      CommentGroup: groups
    }
  }

  async findByCode(code: string): Promise<DbSubject | null> {
    const { rows } = await pool.query<DbSubject>(
      `SELECT * FROM "Subject" WHERE code = $1`,
      [code]
    )
    return rows[0] ?? null
  }

  async create(data: { code: string; title?: string; studiedComment?: string }) {
    const id = data.code
    const { rows } = await pool.query<DbSubject>(
      `INSERT INTO "Subject" (id, code, title, "studiedComment")
       VALUES ($1, $2, $3, $4)
       RETURNING *`,
      [id, data.code, data.title ?? null, data.studiedComment ?? null]
    )
    return { ...rows[0], Class: [], User: [] }
  }

  async update(
    id: string,
    data: { code?: string; title?: string; studiedComment?: string }
  ) {
    const sets: string[] = []
    const params: any[] = []
    let idx = 1

    if (data.code !== undefined) { sets.push(`code = $${idx++}`); params.push(data.code) }
    if (data.title !== undefined) { sets.push(`title = $${idx++}`); params.push(data.title) }
    if (data.studiedComment !== undefined) { sets.push(`"studiedComment" = $${idx++}`); params.push(data.studiedComment) }

    params.push(id)
    const { rows } = await pool.query<DbSubject>(
      `UPDATE "Subject" SET ${sets.join(', ')} WHERE id = $${idx} RETURNING *`,
      params
    )
    return { ...rows[0], Class: [], User: [] }
  }

  async delete(id: string): Promise<void> {
    await pool.query(`DELETE FROM "Subject" WHERE id = $1`, [id])
  }

  async assignUser(subjectId: string, userId: string) {
    await pool.query(
      `INSERT INTO "_SubjectToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [subjectId, userId]
    )
    const { rows: users } = await pool.query(
      `SELECT u.id, u.username FROM "User" u
       JOIN "_SubjectToUser" su ON su."B" = u.id WHERE su."A" = $1`,
      [subjectId]
    )
    return { User: users }
  }

  async removeUser(subjectId: string, userId: string) {
    await pool.query(
      `DELETE FROM "_SubjectToUser" WHERE "A" = $1 AND "B" = $2`,
      [subjectId, userId]
    )
    const { rows: users } = await pool.query(
      `SELECT u.id, u.username FROM "User" u
       JOIN "_SubjectToUser" su ON su."B" = u.id WHERE su."A" = $1`,
      [subjectId]
    )
    return { User: users }
  }
}

export const subjectRepository = new SubjectRepository()
```

**Step 2: Verify TypeScript**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "subject-repository"
```

Expected: No errors.

**Step 3: Commit**

```bash
git add lib/db/repositories/subject-repository.ts
git commit -m "feat: replace Prisma with pg in subject-repository"
```

---

## Task 6: Rewrite `lib/db/repositories/class-repository.ts`

Most complex repository. `findByIdWithAssignments` is a deep nested query — break it into 3 explicit fetches. The `_ClassToUser` junction manages teacher assignments. `_count` select becomes a COUNT query.

**Files:**
- Modify: `lib/db/repositories/class-repository.ts`

**Step 1: Rewrite the file**

```typescript
import { pool } from '@/lib/db'
import { decrypt } from '@/lib/encryption'
import { NotFoundError, ForbiddenError } from '@/lib/errors'
import { createId } from '@paralleldrive/cuid2'
import type { DbClass, DbSubject, DbAssignment, DbPupil, DbPupilCode, DbCommentGroup, DbCommentOption } from '@/lib/types/db'

export class ClassRepository {
  async findById(id: string) {
    const { rows } = await pool.query<DbClass>(
      `SELECT * FROM "Class" WHERE id = $1`,
      [id]
    )

    if (rows.length === 0) {
      throw new NotFoundError(`Class with ID ${id} not found`)
    }

    const cls = rows[0]
    const [{ rows: subjectRows }, { rows: userRows }] = await Promise.all([
      pool.query<DbSubject>(`SELECT * FROM "Subject" WHERE id = $1`, [cls.subjectId]),
      pool.query(
        `SELECT u.id, u.username FROM "User" u
         JOIN "_ClassToUser" cu ON cu."A" = u.id WHERE cu."B" = $1`,
        [id]
      )
    ])

    return { ...cls, Subject: subjectRows[0] ?? null, User: userRows }
  }

  async findByIdWithAssignments(id: string) {
    const cls = await this.findById(id)

    // Fetch subject with comment groups and options
    const { rows: groups } = await pool.query<DbCommentGroup>(
      `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 ORDER BY "displayOrder" ASC`,
      [cls.subjectId]
    )

    const commentGroups = await Promise.all(groups.map(async group => {
      const { rows: options } = await pool.query<DbCommentOption>(
        `SELECT * FROM "CommentOption" WHERE "groupId" = $1 ORDER BY "displayOrder" ASC`,
        [group.id]
      )
      return { ...group, CommentOption: options }
    }))

    // Fetch assignments with active pupils only, ordered by last name
    const { rows: assignments } = await pool.query<DbAssignment & { pupil_admissionNumber: string; pupil_firstName: string; pupil_lastName: string; pupil_gender: string; pupil_isActive: boolean; pupil_form: string | null }>(
      `SELECT a.*, p."admissionNumber" as "pupil_admissionNumber", p."firstName" as "pupil_firstName",
              p."lastName" as "pupil_lastName", p."gender" as "pupil_gender",
              p."isActive" as "pupil_isActive", p."form" as "pupil_form"
       FROM "Assignment" a
       JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
       WHERE a."classId" = $1 AND p."isActive" = true
       ORDER BY p."lastName" ASC`,
      [id]
    )

    // Fetch pupil codes for all assignments
    const assignmentIds = assignments.map(a => a.id)
    let pupilCodes: (DbPupilCode & { CommentGroup: DbCommentGroup })[] = []

    if (assignmentIds.length > 0) {
      const { rows: pcRows } = await pool.query(
        `SELECT pc.*, cg.id as "cg_id", cg.name as "cg_name", cg."displayOrder" as "cg_displayOrder",
                cg."subjectId" as "cg_subjectId", cg.title as "cg_title",
                cg."isLinked" as "cg_isLinked", cg."linkedField" as "cg_linkedField"
         FROM "PupilCode" pc
         JOIN "CommentGroup" cg ON cg.id = pc."groupId"
         WHERE pc."assignmentId" = ANY($1::text[])`,
        [assignmentIds]
      )
      pupilCodes = pcRows.map(row => ({
        id: row.id,
        assignmentId: row.assignmentId,
        groupId: row.groupId,
        code: row.code,
        CommentGroup: {
          id: row.cg_id,
          name: row.cg_name,
          displayOrder: row.cg_displayOrder,
          subjectId: row.cg_subjectId,
          title: row.cg_title,
          isLinked: row.cg_isLinked,
          linkedField: row.cg_linkedField
        }
      }))
    }

    const pcByAssignment = new Map<string, typeof pupilCodes>()
    for (const pc of pupilCodes) {
      const list = pcByAssignment.get(pc.assignmentId) ?? []
      list.push(pc)
      pcByAssignment.set(pc.assignmentId, list)
    }

    const hydratedAssignments = assignments.map(a => ({
      id: a.id,
      pupilId: a.pupilId,
      classId: a.classId,
      eoyLevel: a.eoyLevel,
      targetLevel: a.targetLevel,
      actualLevel: a.actualLevel,
      finalComment: a.finalComment,
      linkedData: a.linkedData,
      checkStatus: a.checkStatus,
      checkNote: a.checkNote,
      checkedAt: a.checkedAt,
      checkedById: a.checkedById,
      Pupil: {
        admissionNumber: a.pupil_admissionNumber,
        firstName: decrypt(a.pupil_firstName),
        lastName: decrypt(a.pupil_lastName),
        gender: a.pupil_gender,
        isActive: a.pupil_isActive,
        form: a.pupil_form
      },
      PupilCode: pcByAssignment.get(a.id) ?? []
    }))

    return {
      ...cls,
      Subject: {
        ...cls.Subject,
        CommentGroup: commentGroups
      },
      Assignment: hydratedAssignments
    }
  }

  async findByTeacherId(teacherId: string) {
    const { rows } = await pool.query<DbClass>(
      `SELECT c.* FROM "Class" c
       JOIN "_ClassToUser" cu ON cu."B" = c.id
       WHERE cu."A" = $1
       ORDER BY c.name ASC`,
      [teacherId]
    )

    return Promise.all(rows.map(async cls => {
      const [{ rows: subjectRows }, { rows: countRows }] = await Promise.all([
        pool.query<DbSubject>(`SELECT * FROM "Subject" WHERE id = $1`, [cls.subjectId]),
        pool.query<{ count: string }>(
          `SELECT COUNT(*) as count FROM "Assignment" WHERE "classId" = $1`,
          [cls.id]
        )
      ])
      return {
        ...cls,
        Subject: subjectRows[0] ?? null,
        _count: { Assignment: parseInt(countRows[0].count) }
      }
    }))
  }

  async assignTeachers(classId: string, teacherIds: string[]) {
    const client = await pool.connect()
    try {
      await client.query('BEGIN')
      await client.query(`DELETE FROM "_ClassToUser" WHERE "B" = $1`, [classId])
      for (const tid of teacherIds) {
        await client.query(
          `INSERT INTO "_ClassToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [tid, classId]
        )
      }
      await client.query('COMMIT')
    } catch (e) {
      await client.query('ROLLBACK')
      throw e
    } finally {
      client.release()
    }

    const { rows: userRows } = await pool.query(
      `SELECT u.* FROM "User" u JOIN "_ClassToUser" cu ON cu."A" = u.id WHERE cu."B" = $1`,
      [classId]
    )
    return { User: userRows }
  }

  async authorizeAccess(classId: string, userId: string, userRoles: string[]): Promise<boolean> {
    if (userRoles.includes('admin')) return true

    const cls = await this.findById(classId)

    if (userRoles.includes('hod')) {
      const { rows } = await pool.query(
        `SELECT 1 FROM "_SubjectToUser" WHERE "A" = $1 AND "B" = $2`,
        [cls.subjectId, userId]
      )
      if (rows.length > 0) return true
    }

    return cls.User.some((t: any) => t.id === userId)
  }

  async requireAccess(classId: string, userId: string, userRoles: string[]): Promise<void> {
    const hasAccess = await this.authorizeAccess(classId, userId, userRoles)
    if (!hasAccess) {
      throw new ForbiddenError('You do not have access to this class')
    }
  }

  async findAll() {
    const { rows } = await pool.query<DbClass>(`SELECT * FROM "Class" ORDER BY name ASC`)

    return Promise.all(rows.map(async cls => {
      const [{ rows: subjectRows }, { rows: userRows }, { rows: countRows }] = await Promise.all([
        pool.query(`SELECT id, code, title FROM "Subject" WHERE id = $1`, [cls.subjectId]),
        pool.query(
          `SELECT u.id, u.username FROM "User" u
           JOIN "_ClassToUser" cu ON cu."A" = u.id WHERE cu."B" = $1`,
          [cls.id]
        ),
        pool.query<{ count: string }>(
          `SELECT COUNT(*) as count FROM "Assignment" WHERE "classId" = $1`,
          [cls.id]
        )
      ])

      return {
        ...cls,
        Subject: subjectRows[0] ?? null,
        User: userRows,
        _count: { Assignment: parseInt(countRows[0].count) }
      }
    }))
  }

  async create(data: { name: string; year: string | null; subjectId: string }) {
    const { rows } = await pool.query<DbClass>(
      `INSERT INTO "Class" (id, name, year, "subjectId") VALUES ($1, $2, $3, $4) RETURNING *`,
      [createId(), data.name, data.year, data.subjectId]
    )
    return rows[0]
  }

  async update(classId: string, data: { name?: string; year?: string | null; subjectId?: string }) {
    const sets: string[] = []
    const params: any[] = []
    let idx = 1

    if (data.name !== undefined) { sets.push(`name = $${idx++}`); params.push(data.name) }
    if (data.year !== undefined) { sets.push(`year = $${idx++}`); params.push(data.year) }
    if (data.subjectId !== undefined) { sets.push(`"subjectId" = $${idx++}`); params.push(data.subjectId) }

    params.push(classId)
    const { rows } = await pool.query<DbClass>(
      `UPDATE "Class" SET ${sets.join(', ')} WHERE id = $${idx} RETURNING *`,
      params
    )

    const [{ rows: subjectRows }, { rows: userRows }] = await Promise.all([
      pool.query(`SELECT id, code, title FROM "Subject" WHERE id = $1`, [rows[0].subjectId]),
      pool.query(
        `SELECT u.id, u.username FROM "User" u
         JOIN "_ClassToUser" cu ON cu."A" = u.id WHERE cu."B" = $1`,
        [classId]
      )
    ])

    return { ...rows[0], Subject: subjectRows[0] ?? null, User: userRows }
  }

  async delete(classId: string): Promise<void> {
    await pool.query(`DELETE FROM "Class" WHERE id = $1`, [classId])
  }

  async assignPupils(classId: string, pupilAdmissionNumbers: string[]) {
    if (pupilAdmissionNumbers.length === 0) return { count: 0 }

    const values = pupilAdmissionNumbers.map((_, i) => `($${i * 3 + 1}, $${i * 3 + 2}, $${i * 3 + 3})`).join(', ')
    const params = pupilAdmissionNumbers.flatMap(admNo => [
      `${classId}-${admNo}`,
      admNo,
      classId
    ])

    const { rowCount } = await pool.query(
      `INSERT INTO "Assignment" (id, "pupilId", "classId")
       VALUES ${values}
       ON CONFLICT DO NOTHING`,
      params
    )
    return { count: rowCount ?? 0 }
  }

  async getPupils(classId: string) {
    const { rows } = await pool.query(
      `SELECT p.*, a.id as "assignmentId"
       FROM "Assignment" a
       JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
       WHERE a."classId" = $1
       ORDER BY p."lastName" ASC`,
      [classId]
    )

    return rows.map(row => ({
      admissionNumber: row.admissionNumber,
      firstName: decrypt(row.firstName),
      lastName: decrypt(row.lastName),
      gender: row.gender,
      isActive: row.isActive,
      form: row.form,
      assignmentId: row.assignmentId
    }))
  }

  async removePupil(classId: string, pupilAdmissionNumber: string) {
    const { rowCount } = await pool.query(
      `DELETE FROM "Assignment" WHERE "classId" = $1 AND "pupilId" = $2`,
      [classId, pupilAdmissionNumber]
    )
    return { count: rowCount ?? 0 }
  }

  async removePupils(classId: string, pupilAdmissionNumbers: string[]) {
    const { rowCount } = await pool.query(
      `DELETE FROM "Assignment" WHERE "classId" = $1 AND "pupilId" = ANY($2::text[])`,
      [classId, pupilAdmissionNumbers]
    )
    return { count: rowCount ?? 0 }
  }
}

export const classRepository = new ClassRepository()
```

**Step 2: Verify TypeScript**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "class-repository"
```

Expected: No errors.

**Step 3: Commit**

```bash
git add lib/db/repositories/class-repository.ts
git commit -m "feat: replace Prisma with pg in class-repository"
```

---

## Task 7: Update `lib/audit-log.ts`

One `prisma.auditLog.create` call. Replace with a direct INSERT.

**Files:**
- Modify: `lib/audit-log.ts`

**Step 1: Replace the import and the create call**

Change the import at the top from:
```typescript
import { prisma } from '@/lib/prisma'
```
to:
```typescript
import { pool } from '@/lib/db'
```

Replace the `prisma.auditLog.create` block (lines 139–151) with:
```typescript
    await pool.query(
      `INSERT INTO "AuditLog" (id, "userId", username, action, "entityType", "entityId", details, "ipAddress", "userAgent")
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)`,
      [
        createId(),
        userId ?? null,
        username ?? null,
        params.action,
        params.entityType ?? null,
        params.entityId ?? null,
        params.details ? encrypt(JSON.stringify(params.details)) : null,
        ipAddress ?? null,
        userAgent ?? null
      ]
    )
```

**Step 2: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "audit-log"
```

**Step 3: Commit**

```bash
git add lib/audit-log.ts
git commit -m "feat: replace Prisma with pg in audit-log"
```

---

## Task 8: Update `lib/auth/with-role.ts`

One `prisma.user.findUnique` call to check if user is still active.

**Files:**
- Modify: `lib/auth/with-role.ts`

**Step 1: Replace import and query**

Change import from `import { prisma } from '@/lib/prisma'` to `import { pool } from '@/lib/db'`.

Replace the `prisma.user.findUnique` block with:
```typescript
    const { rows } = await pool.query<{ isActive: boolean }>(
      `SELECT "isActive" FROM "User" WHERE id = $1`,
      [session.user.id]
    )
    const user = rows[0] ?? null
```

**Step 2: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "with-role"
```

**Step 3: Commit**

```bash
git add lib/auth/with-role.ts
git commit -m "feat: replace Prisma with pg in with-role auth middleware"
```

---

## Task 9: Update `lib/server-actions/create-user.ts`

Two Prisma calls: `findUnique` for duplicate check, `create` for user creation.

**Files:**
- Modify: `lib/server-actions/create-user.ts`

**Step 1: Replace import and both queries**

Change import to `import { pool } from '@/lib/db'`.

Replace `prisma.user.findUnique` with:
```typescript
    const { rows: existingRows } = await pool.query(
      `SELECT id FROM "User" WHERE username = $1`,
      [username]
    )
    if (existingRows.length > 0) {
      return { success: false, error: "Username already exists" }
    }
```

Replace `prisma.user.create` with:
```typescript
    const userId = createId()
    await pool.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)`,
      [userId, username, hashedPassword]
    )

    if (role) {
      const { rows: roleRows } = await pool.query(
        `SELECT id FROM "Role" WHERE name = $1`,
        [role]
      )
      if (roleRows.length > 0) {
        await pool.query(
          `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [roleRows[0].id, userId]
        )
      }
    }

    const newUser = { id: userId }
```

**Step 2: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "create-user"
```

**Step 3: Commit**

```bash
git add lib/server-actions/create-user.ts
git commit -m "feat: replace Prisma with pg in create-user server action"
```

---

## Task 10: Update `lib/server-actions/admin.ts`

24 direct Prisma calls. Most are simple `findUnique` / `update` / `delete` for audit-log context lookups. Replace each pattern.

**Files:**
- Modify: `lib/server-actions/admin.ts`

**Step 1: Change the import**

Replace:
```typescript
import { prisma } from '@/lib/prisma'
```
with:
```typescript
import { pool } from '@/lib/db'
```

**Step 2: Replace each `prisma.*` call**

Work through the file top to bottom. For each pattern:

**`prisma.user.findUnique({ where: { id }, include/select })`** →
```typescript
const { rows } = await pool.query(`SELECT * FROM "User" WHERE id = $1`, [id])
// then if roles needed: also query _RoleToUser JOIN Role
```

**`prisma.user.update({ where: { id }, data: { isActive } })`** →
```typescript
await pool.query(`UPDATE "User" SET "isActive" = $1 WHERE id = $2`, [isActive, id])
```

**`prisma.subject.findUnique({ where: { id } })`** →
```typescript
const { rows } = await pool.query(`SELECT * FROM "Subject" WHERE id = $1`, [id])
```

**`prisma.class.findUnique({ where: { id }, include: { Subject, User } })`** →
```typescript
const { rows } = await pool.query(`SELECT * FROM "Class" WHERE id = $1`, [id])
// plus subject query and _ClassToUser JOIN for users
```

**`prisma.user.findMany({ where: { id: { in: teacherIds } }, select: { id, username } })`** →
```typescript
const { rows } = await pool.query(
  `SELECT id, username FROM "User" WHERE id = ANY($1::text[])`,
  [teacherIds]
)
```

**`prisma.pupil.groupBy({ by: ['form'], where: { form: { not: null }, isActive: true } })`** →
```typescript
const { rows } = await pool.query(
  `SELECT DISTINCT form FROM "Pupil" WHERE form IS NOT NULL AND "isActive" = true ORDER BY form ASC`
)
const forms = rows.map(r => r.form)
```

**`prisma.deadline.findMany()`** →
```typescript
const { rows } = await pool.query(`SELECT * FROM "Deadline" ORDER BY date ASC`)
```

**`prisma.deadline.create({ data: { id, title, date, description } })`** →
```typescript
await pool.query(
  `INSERT INTO "Deadline" (id, title, date, description) VALUES ($1, $2, $3, $4)`,
  [createId(), title, new Date(date), description ?? null]
)
```

**`prisma.deadline.update` / `prisma.deadline.delete`** → equivalent UPDATE/DELETE statements.

**`prisma.user.findMany({ select: { id, username, Role } })`** (getAllUsers) →
```typescript
const { rows: users } = await pool.query(`SELECT * FROM "User" ORDER BY username ASC`)
// then for each user attach roles, or do a JOIN:
const { rows } = await pool.query(`
  SELECT u.id, u.username, r.name as "roleName"
  FROM "User" u
  LEFT JOIN "_RoleToUser" ru ON ru."B" = u.id
  LEFT JOIN "Role" r ON r.id = ru."A"
  ORDER BY u.username ASC
`)
// group by user in JS
```

**Step 3: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "admin.ts"
```

**Step 4: Commit**

```bash
git add lib/server-actions/admin.ts
git commit -m "feat: replace Prisma with pg in admin server actions"
```

---

## Task 11: Update `lib/server-actions/hod.ts`

28 direct Prisma calls. Patterns: class CRUD, comment group CRUD with `aggregate` for max `displayOrder`, comment option CRUD, subject update.

**Files:**
- Modify: `lib/server-actions/hod.ts`

**Step 1: Change import**

`import { pool } from '@/lib/db'`

**Step 2: Replace each call**

**`prisma.subject.findFirst({ where: { id, User: { some: { id } } } })`** (the HOD access check) →
```typescript
const { rows } = await pool.query(
  `SELECT 1 FROM "Subject" s
   JOIN "_SubjectToUser" su ON su."A" = s.id
   WHERE s.id = $1 AND su."B" = $2`,
  [subjectId, session.user.id]
)
```

**`prisma.class.create`** →
```typescript
const { rows } = await pool.query(
  `INSERT INTO "Class" (id, name, year, "subjectId") VALUES ($1, $2, $3, $4) RETURNING *`,
  [createId(), name, year, subjectId]
)
```

**`prisma.commentGroup.aggregate({ where, _max: { displayOrder } })`** →
```typescript
const { rows } = await pool.query<{ max: number | null }>(
  `SELECT MAX("displayOrder") as max FROM "CommentGroup" WHERE "subjectId" = $1`,
  [subjectId]
)
const nextOrder = (rows[0].max ?? -1) + 1
```

**`prisma.commentGroup.create`** →
```typescript
await pool.query(
  `INSERT INTO "CommentGroup" (id, name, title, "subjectId", "displayOrder", "isLinked", "linkedField")
   VALUES ($1, $2, $3, $4, $5, $6, $7)`,
  [createId(), name, title, subjectId, nextOrder, isLinked, linkedField]
)
```

**`prisma.commentGroup.update({ where, data: { displayOrder } })`** (reorder) →
```typescript
await pool.query(
  `UPDATE "CommentGroup" SET "displayOrder" = $1 WHERE id = $2`,
  [item.order, item.id]
)
```

**`prisma.commentGroup.findUnique({ where, select: { subjectId } })`** →
```typescript
const { rows } = await pool.query(
  `SELECT "subjectId" FROM "CommentGroup" WHERE id = $1`,
  [groupId]
)
```

**`prisma.commentOption.aggregate`** → same pattern as comment group max.

**`prisma.subject.update({ where, data: { commentFormat } })`** →
```typescript
await pool.query(
  `UPDATE "Subject" SET "commentFormat" = $1 WHERE id = $2`,
  [commentFormat, subjectId]
)
```

**Assignment CRUD** (`prisma.assignment.create/update/delete`) → direct INSERT/UPDATE/DELETE.

**Step 3: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "hod.ts"
```

**Step 4: Commit**

```bash
git add lib/server-actions/hod.ts
git commit -m "feat: replace Prisma with pg in hod server actions"
```

---

## Task 12: Update `lib/server-actions/ccg.ts`

This file uses `(prisma as any).commonCommentGroup.*` and `(prisma as any).commonCommentOption.*` — the type casting is because these models were added after the main schema was set up. Replace all with `pool.query()` against `"CommonCommentGroup"` and `"CommonCommentOption"` tables.

**Files:**
- Modify: `lib/server-actions/ccg.ts`

**Step 1: Read the full file first**

```bash
cat lib/server-actions/ccg.ts
```

Then replace `import { prisma } from '@/lib/prisma'` with `import { pool } from '@/lib/db'` and translate each `(prisma as any).*` call following the same patterns as Tasks 10 and 11.

**Step 2: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "ccg.ts"
```

**Step 3: Commit**

```bash
git add lib/server-actions/ccg.ts
git commit -m "feat: replace Prisma with pg in ccg server actions"
```

---

## Task 13: Update `lib/server-actions/comment-check.ts`

7 Prisma calls. Key ones: `assignment.findUnique` with nested `Class.Subject` include, `assignment.update`, `assignment.groupBy` for stats.

**Files:**
- Modify: `lib/server-actions/comment-check.ts`

**Step 1: Change import**

`import { pool } from '@/lib/db'`

**Step 2: Replace calls**

**`prisma.assignment.findUnique({ include: { Class: { include: { Subject } } } })`** →
```typescript
const { rows: aRows } = await pool.query(
  `SELECT a.*, c.id as "class_id", c.name as "class_name", c."subjectId" as "class_subjectId",
          s.id as "subject_id", s.code as "subject_code"
   FROM "Assignment" a
   JOIN "Class" c ON c.id = a."classId"
   LEFT JOIN "Subject" s ON s.id = c."subjectId"
   WHERE a.id = $1`,
  [assignmentId]
)
// then reshape into { ...assignment, Class: { ...class, Subject: subject } }
```

**`prisma.assignment.findMany` with deep include for review list** → multi-query approach (fetch assignments, then pupils, then pupil codes with comment groups).

**`prisma.assignment.groupBy({ by: ['checkStatus'], where: { Class: { subjectId } }, _count: true })`** →
```typescript
const { rows } = await pool.query(
  `SELECT "checkStatus", COUNT(*) as count
   FROM "Assignment" a
   JOIN "Class" c ON c.id = a."classId"
   WHERE c."subjectId" = $1
   GROUP BY "checkStatus"`,
  [subjectId]
)
```

**Step 3: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "comment-check"
```

**Step 4: Commit**

```bash
git add lib/server-actions/comment-check.ts
git commit -m "feat: replace Prisma with pg in comment-check server actions"
```

---

## Task 14: Update `lib/server-actions/linked-data.ts`

3 calls: one raw SQL already (`prisma.$queryRaw` — just becomes `pool.query`), one `findMany`, one `update` in a loop.

**Files:**
- Modify: `lib/server-actions/linked-data.ts`

**Step 1: Change import and replace calls**

`prisma.$queryRaw\`...\`` → `pool.query('...')` (the SQL is already written — just move it into pool.query)

`prisma.assignment.findMany({ where: { classId }, select: { id, pupilId } })` →
```typescript
const { rows: assignments } = await pool.query(
  `SELECT id, "pupilId" FROM "Assignment" WHERE "classId" = $1`,
  [classId]
)
```

`prisma.assignment.update({ where: { id }, data: { linkedData } })` →
```typescript
await pool.query(
  `UPDATE "Assignment" SET "linkedData" = $1 WHERE id = $2`,
  [JSON.stringify(linkedData), assignmentId]
)
```

**Step 2: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "linked-data"
```

**Step 3: Commit**

```bash
git add lib/server-actions/linked-data.ts
git commit -m "feat: replace Prisma with pg in linked-data server actions"
```

---

## Task 15: Update `lib/server-actions/audit-log.ts`

9 calls — mostly `findMany`/`count`/`groupBy` on `AuditLog`.

**Files:**
- Modify: `lib/server-actions/audit-log.ts`

**Step 1: Change import**

`import { pool } from '@/lib/db'`

**Step 2: Replace `getAuditLogs`**

Build WHERE clause dynamically in SQL. Then:
```typescript
const { rows: logs } = await pool.query(
  `SELECT * FROM "AuditLog" ${whereClause} ORDER BY "createdAt" DESC LIMIT $N OFFSET $M`,
  params
)
const { rows: countRows } = await pool.query(
  `SELECT COUNT(*) as count FROM "AuditLog" ${whereClause}`,
  countParams
)
```

**Step 3: Replace `getAuditLogActions` (`groupBy: ['action']`)**

```typescript
const { rows } = await pool.query(
  `SELECT DISTINCT action FROM "AuditLog" ORDER BY action ASC`
)
return { success: true as const, actions: rows.map(r => r.action) }
```

**Step 4: Replace `getAuditLogEntityTypes`**

```typescript
const { rows } = await pool.query(
  `SELECT DISTINCT "entityType" FROM "AuditLog" WHERE "entityType" IS NOT NULL ORDER BY "entityType" ASC`
)
```

**Step 5: Replace `getAuditLogStats` (multiple `count` calls)**

```typescript
const { rows } = await pool.query(`
  SELECT
    COUNT(*) FILTER (WHERE true) as total,
    COUNT(*) FILTER (WHERE "createdAt" >= $1) as today,
    COUNT(*) FILTER (WHERE "createdAt" >= $2) as week,
    COUNT(*) FILTER (WHERE action = 'sign_in' AND "createdAt" >= $1) as sign_ins,
    COUNT(*) FILTER (WHERE action NOT IN ('sign_in','sign_out','sign_in_failed') AND "createdAt" >= $1) as changes
  FROM "AuditLog"
`, [today, weekAgo])
```

**Step 6: Verify**

```bash
npx tsc --noEmit --incremental false 2>&1 | grep "audit-log"
```

**Step 7: Commit**

```bash
git add lib/server-actions/audit-log.ts
git commit -m "feat: replace Prisma with pg in audit-log server actions"
```

---

## Task 16: Final cleanup — remove Prisma

**Files:**
- Modify: `package.json`
- Delete: `lib/prisma.ts`
- Delete: `prisma.config.ts`

**Step 1: Update the build script in `package.json`**

Change:
```json
"build": "prisma generate && next build",
```
to:
```json
"build": "next build",
```

**Step 2: Uninstall Prisma packages**

```bash
npm uninstall @prisma/client @prisma/adapter-pg prisma
```

**Step 3: Delete `lib/prisma.ts` and `prisma.config.ts`**

```bash
rm lib/prisma.ts prisma.config.ts
```

**Step 4: Final TypeScript check**

```bash
npx tsc --noEmit --incremental false 2>&1
```

Expected: Zero errors. If any remain, they point to a missed `prisma.*` reference — fix them now.

**Step 5: Verify the app builds**

```bash
npm run build
```

Expected: Clean build, no Prisma generate step.

**Step 6: Final commit**

```bash
git add package.json package-lock.json
git rm lib/prisma.ts prisma.config.ts
git commit -m "chore: remove Prisma packages and config — now using pg directly"
```

---

## Verification

After all tasks are complete:

1. Run unit tests: `npm test`
2. Start dev server: `npm run dev`
3. Smoke-test critical paths: sign in, load a class, save a comment, admin user management
4. Check audit log entries are still being written
