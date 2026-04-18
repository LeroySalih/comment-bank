# E2E Test Suite Setup — Design Spec

**Date:** 2026-04-18  
**Status:** Approved

## Goal

Fix the broken test infrastructure so the existing Playwright e2e suite can run. No new tests are being added — the suite already exists in `tests/`.

## Problem

Two things prevent `npm run test:e2e` from working:

1. `.env.test` does not exist — Playwright's `webServer` command calls `dotenv -e .env.test` which fails immediately.
2. `scripts/setup-test-db.sh` calls `npx prisma migrate reset` — Prisma CLI has been removed from the project.

## Files Changed

| File | Action |
|---|---|
| `.env.test` | Create |
| `scripts/setup-test-db.sh` | Rewrite |

## `.env.test`

Mirrors `.env` with three differences:

- `ENV=TEST`
- `DATABASE_URL` points to `comment_bank_test` (same host, port, user, password as dev)
- `NEXTAUTH_URL=http://localhost:3001` (matches Playwright `baseURL` and `webServer.port`)

```env
ENV=TEST
NEXT_PUBLIC_ENV=TEST
DATABASE_URL="postgresql://postgres:your-super-secret-and-long-postgres-password@localhost:5432/comment_bank_test"
NEXTAUTH_SECRET="supersecret123"
NEXTAUTH_URL="http://localhost:3001"
NEXT_PUBLIC_SITE_URL="http://localhost:3001"
PUPIL_ENCRYPTION_KEY=800b64514e68b86e3985f30c8896e4952ee2e9b111377025f03031a05cfa2486
```

## `scripts/setup-test-db.sh` Rewrite

### Steps

1. **Load `.env.test`** — source the file; abort if it does not exist.
2. **Safety check** — abort if `DATABASE_URL` does not contain `comment_bank_test` (unchanged from current script).
3. **Drop and recreate the test database** — via `docker exec postgres17 psql -U postgres`:
   ```sql
   DROP DATABASE IF EXISTS comment_bank_test;
   CREATE DATABASE comment_bank_test;
   ```
4. **Apply migrations** — call `migrations/migrate.sh postgres17 comment_bank_test postgres`.
5. **Seed test data** — run `npx tsx prisma/seed.ts` with `DATABASE_URL` set to the test URL.

### Unchanged

- The `npm run test:e2e:setup` entry point (`./scripts/setup-test-db.sh && npx playwright test`) stays the same.
- `npm run test:e2e` (runs Playwright without setup) also unchanged.
- `migrations/migrate.sh` is not modified.
- `prisma/seed.ts` is not modified.

## Docker Assumption

`migrations/migrate.sh` assumes a Docker container named `postgres17` is running. `setup-test-db.sh` makes the same assumption. If the container name differs, pass it as the first argument to `migrate.sh` or set it as a variable at the top of the script.

## Test Credentials

Defined in `tests/helpers.ts` (`TEST_USERS`) and seeded by `prisma/seed.ts`:

| Username | Password | Role |
|---|---|---|
| `admin` | `password` | admin |
| `leroysalih` | `password` | hod |
| `teacher` | `password` | teacher |
| `teacher2` | `password` | teacher |
| `teacher3` | `password` | teacher |
