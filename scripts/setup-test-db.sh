#!/usr/bin/env bash
set -euo pipefail

# Setup test database environment
# Usage: ./scripts/setup-test-db.sh
#
# This script:
# 1. Loads env vars from .env.test
# 2. Validates DATABASE_URL points to the test database (safety check)
# 3. Drops and recreates the test database via prisma migrate reset
# 4. Seeds with test data from prisma/seed.ts

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

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

# Drop, recreate, and apply all migrations
echo "==> Resetting test database (drop + migrate)..."
cd "$PROJECT_DIR"
if ! PRISMA_USER_CONSENT_FOR_DANGEROUS_AI_ACTION="yes" npx prisma migrate reset --force --skip-generate --skip-seed; then
  echo ""
  echo "FAILED: prisma migrate reset failed."
  exit 1
fi

# Run seed explicitly
echo ""
echo "==> Seeding test database..."
if ! npx tsx prisma/seed.ts; then
  echo ""
  echo "FAILED: seed script failed."
  exit 1
fi

echo ""
echo "==> Test database ready."
echo "    - Migrations applied"
echo "    - Test data seeded"
echo "    - Database: comment_bank_test"
