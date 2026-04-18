# E2E Test Suite Setup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the broken e2e test infrastructure so `npm run test:e2e` can run the existing Playwright suite against a dedicated test database.

**Architecture:** Create `.env.test` with test-database credentials, then rewrite `scripts/setup-test-db.sh` to drop/recreate `comment_bank_test` via Docker, apply migrations via the existing `migrations/migrate.sh`, and seed via `prisma/seed.ts`.

**Tech Stack:** Bash, Docker (`postgres17` container), `migrations/migrate.sh`, `npx tsx`, Playwright

---

## File Map

| File | Action | Responsibility |
|---|---|---|
| `.env.test` | Create | Test environment variables — points app at `comment_bank_test` on port 3001 |
| `scripts/setup-test-db.sh` | Rewrite | Drop/recreate test DB, apply migrations, seed test data |

---

### Task 1: Create `.env.test`

**Files:**
- Create: `.env.test`

- [ ] **Step 1: Create the file**

Create `.env.test` at the project root with these exact contents:

```env
ENV=TEST
NEXT_PUBLIC_ENV=TEST
DATABASE_URL="postgresql://postgres:your-super-secret-and-long-postgres-password@localhost:5432/comment_bank_test"
NEXTAUTH_SECRET="supersecret123"
NEXTAUTH_URL="http://localhost:3001"
NEXT_PUBLIC_SITE_URL="http://localhost:3001"
PUPIL_ENCRYPTION_KEY=800b64514e68b86e3985f30c8896e4952ee2e9b111377025f03031a05cfa2486
```

> The password and encryption key are copied verbatim from `.env`. The only differences from `.env` are `ENV=TEST`, the database name (`comment_bank_test`), and `NEXTAUTH_URL` pointing to port 3001.

- [ ] **Step 2: Verify the file is gitignored**

Run:
```bash
git check-ignore -v .env.test
```

Expected output should match `.gitignore` (e.g. `.gitignore:N:.env*`). If it is **not** ignored, add `.env.test` to `.gitignore` before committing.

- [ ] **Step 3: Commit**

```bash
git commit -m "chore: add .env.test for e2e test environment"
```

> `.env.test` should be gitignored and therefore will not appear in `git status`. If it does appear (not ignored), add it to `.gitignore` and commit that instead.

---

### Task 2: Rewrite `scripts/setup-test-db.sh`

**Files:**
- Modify: `scripts/setup-test-db.sh`

- [ ] **Step 1: Read the current file**

Read `scripts/setup-test-db.sh` to confirm the current content before overwriting.

- [ ] **Step 2: Replace the file contents**

Replace the entire file with:

```bash
#!/usr/bin/env bash
set -euo pipefail

# Setup test database environment
# Usage: ./scripts/setup-test-db.sh
#
# 1. Loads env vars from .env.test
# 2. Validates DATABASE_URL points to comment_bank_test (safety check)
# 3. Drops and recreates comment_bank_test via Docker
# 4. Applies all SQL migrations via migrations/migrate.sh
# 5. Seeds test data via prisma/seed.ts

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

CONTAINER="postgres17"
DB_NAME="comment_bank_test"
DB_USER="postgres"

echo "==> Setting up test database..."

# Load test env
if [ ! -f "$PROJECT_DIR/.env.test" ]; then
  echo "Error: .env.test not found in $PROJECT_DIR"
  exit 1
fi

set -a
source "$PROJECT_DIR/.env.test"
set +a

# Safety check: ensure DATABASE_URL targets the test database
if [[ "$DATABASE_URL" != *"comment_bank_test"* ]]; then
  echo "ABORT: DATABASE_URL does not point to comment_bank_test."
  echo "       Current value: $DATABASE_URL"
  echo "       This script refuses to run against a non-test database."
  exit 1
fi

echo "==> Using DATABASE_URL pointing to comment_bank_test"

# Drop and recreate the test database
echo "==> Dropping and recreating $DB_NAME..."
docker exec "$CONTAINER" psql -U "$DB_USER" -c "DROP DATABASE IF EXISTS $DB_NAME;"
docker exec "$CONTAINER" psql -U "$DB_USER" -c "CREATE DATABASE $DB_NAME;"

# Apply all migrations
echo "==> Applying migrations..."
"$PROJECT_DIR/migrations/migrate.sh" "$CONTAINER" "$DB_NAME" "$DB_USER"

# Seed test data
echo "==> Seeding test database..."
cd "$PROJECT_DIR"
if ! DATABASE_URL="$DATABASE_URL" npx tsx prisma/seed.ts; then
  echo ""
  echo "FAILED: seed script failed."
  exit 1
fi

echo ""
echo "==> Test database ready."
echo "    - Migrations applied"
echo "    - Test data seeded"
echo "    - Database: $DB_NAME"
```

- [ ] **Step 3: Ensure the script is executable**

```bash
chmod +x scripts/setup-test-db.sh
```

- [ ] **Step 4: Commit**

```bash
git add scripts/setup-test-db.sh
git commit -m "chore: rewrite setup-test-db.sh to use migrate.sh instead of prisma"
```

---

### Task 3: Verify the setup runs end-to-end

**Files:** none (verification only)

- [ ] **Step 1: Confirm Docker container is running**

```bash
docker inspect -f '{{.State.Running}}' postgres17
```

Expected: `true`

If the container is not running, start it before proceeding.

- [ ] **Step 2: Run the setup script**

```bash
npm run db:test:reset
```

Expected output (last few lines):
```
==> Test database ready.
    - Migrations applied
    - Test data seeded
    - Database: comment_bank_test
```

The script must exit 0. If it fails, fix the error before moving on.

- [ ] **Step 3: Run the single teacher sign-in test**

```bash
npx playwright test --project=chromium tests/role-teacher.spec.ts -g "Can sign in and see dashboard"
```

Expected: 1 test passes. The Playwright `webServer` block will start the Next.js dev server on port 3001 using `.env.test` automatically.

- [ ] **Step 4: Run the full e2e suite**

```bash
npm run test:e2e
```

Expected: all tests pass (or known-failing tests fail for documented reasons unrelated to setup).

- [ ] **Step 5: Commit any fixups found during verification**

If any small corrections were needed (wrong password, wrong port, etc.), commit them now:

```bash
git add -p
git commit -m "fix: correct .env.test values after verification"
```
