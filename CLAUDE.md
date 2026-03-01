# CLAUDE.md — Project Instructions

## Database Access

**Do not use Prisma.** Prisma has been removed from this project. Use direct `pg` Pool queries instead.

- Import the pool singleton: `import { pool } from '@/lib/db'`
- Use parameterised queries: `pool.query('SELECT ... WHERE id = $1', [id])`
- TypeScript interfaces for all tables are in `lib/types/db.ts` (`DbUser`, `DbClass`, etc.)
- Use `pool.query<DbUser>(...)` generics for type-safe results
- For multi-step writes, use a transaction: `const client = await pool.connect()` with `BEGIN`/`COMMIT`/`ROLLBACK`/`client.release()`
- **Always double-quote SQL aliases that contain uppercase letters** — PostgreSQL folds unquoted aliases to lowercase, so `p."firstName" as pupil_firstName` returns `pupil_firstname` in the result. Use `p."firstName" as "pupil_firstName"` instead.

The `prisma/schema.prisma` file is kept as a table-structure reference only — it is not used at runtime.

## Database Migrations

Prisma CLI is gone, so migrations are plain SQL files. To add a schema change:

1. Create a directory under `migrations/` named `YYYYMMDDHHMMSS_description/`
2. Write a `migration.sql` file inside it with the raw SQL (use `IF NOT EXISTS` / `IF EXISTS` guards where appropriate)
3. Apply it to the database: `psql $DATABASE_URL -f migrations/<dir>/migration.sql`

Also update `prisma/schema.prisma` to keep it in sync as a reference.

## Worktrees

New feature branches should be created as worktrees under `.worktrees/` (already in `.gitignore`).
