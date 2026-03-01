# Design: Remove Prisma, use pg directly

**Date:** 2026-03-01
**Status:** Approved

## Goal

Replace Prisma's generated ORM client with direct `pg` Pool queries across the codebase.
Motivation: more control over SQL. The `pg` package is already installed (used by the current `PrismaPg` adapter).

## Non-Goals

- No change to server-action signatures or React components
- No change to the repository class structure
- No introduction of a query builder (e.g., Kysely)

## Approach

**Keep repositories, swap internals.** All 4 repository classes (`ClassRepository`, `UserRepository`, `SubjectRepository`, `PupilRepository`) stay. Their method signatures and return shapes stay identical. Internally, `prisma.*` calls are replaced with `pool.query(sql, [params])`.

Direct `prisma.*` calls in server-actions, auth middleware, and audit-log are replaced the same way.

---

## Phases

### Phase 1 — Create a `pg` pool singleton

- Replace `lib/prisma.ts` with a `lib/db.ts` that exports a `Pool` instance
- Update all `import { prisma } from '@/lib/prisma'` → `import { pool } from '@/lib/db'`

### Phase 2 — Rewrite the 4 repository files

| File | Prisma call count |
|------|-------------------|
| `lib/db/repositories/class-repository.ts` | 13 |
| `lib/db/repositories/user-repository.ts` | 8 |
| `lib/db/repositories/subject-repository.ts` | 8 |
| `lib/db/repositories/pupil-repository.ts` | 6 |

Key translation patterns:
- `findUnique` / `findMany` → `SELECT ... WHERE ...`
- `create` / `update` / `delete` → `INSERT` / `UPDATE` / `DELETE`
- `include` with relations → secondary queries or JOINs, results restructured in JS
- `_count` selects → `COUNT(*)` subqueries
- `createMany({ skipDuplicates: true })` → `INSERT ... ON CONFLICT DO NOTHING`
- `upsert` → `INSERT ... ON CONFLICT DO UPDATE`
- Many-to-many set operations (User↔Role, User↔Class, User↔Subject) → DELETE + INSERT on junction tables
- `orderBy: { Pupil: { lastName } }` → `ORDER BY p.last_name ASC` in JOIN

### Phase 3 — Fix direct Prisma calls outside repositories

| File | Prisma call count |
|------|-------------------|
| `lib/server-actions/admin.ts` | 24 |
| `lib/server-actions/hod.ts` | 28 |
| `lib/server-actions/audit-log.ts` | 9 |
| `lib/server-actions/comment-check.ts` | 7 |
| `lib/server-actions/linked-data.ts` | 3 |
| `lib/server-actions/create-user.ts` | 2 |
| `lib/auth/with-role.ts` | 3 |

Most of these are simple `findUnique` or `findMany` lookups for audit-log context — straightforward `SELECT` replacements.

### Phase 4 — Define shared TypeScript interfaces

Create `lib/types/db.ts` with interfaces matching each table (replacing Prisma-generated types).
`pool.query<T>()` is generic — callers cast to these types.

### Phase 5 — Remove Prisma

- Uninstall `@prisma/client`, `prisma`, `@prisma/adapter-pg`
- Remove `prisma.config.ts`
- Remove `prisma generate` from `package.json` scripts
- Keep `prisma/schema.prisma` as a table-structure reference (do not delete)

---

## Known Risks

| Risk | Mitigation |
|------|------------|
| Implicit junction table names (`_ClassToUser` etc.) | Confirm actual names with `\dt` in psql before writing queries |
| Deep nested `include` in `findByIdWithAssignments` | Break into 2–3 explicit queries, restructure in JS |
| `groupBy` (used once in `getAllForms`) | Replace with `SELECT DISTINCT form FROM "Pupil" WHERE ...` |
| Type safety loss | Define `lib/types/db.ts` interfaces; use `pool.query<T>()` generics |
| `@default(now())` / `createdAt` | Already in DB schema; no change needed |
| `id` generation | Already uses `createId()` from `@paralleldrive/cuid2` in app code |
