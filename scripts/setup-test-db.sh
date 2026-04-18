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
# Pre-flight: verify container is running
if ! docker inspect -f '{{.State.Running}}' "$CONTAINER" 2>/dev/null | grep -q true; then
  echo "ABORT: Docker container '$CONTAINER' is not running."
  exit 1
fi

echo "==> Dropping and recreating $DB_NAME..."
if ! docker exec "$CONTAINER" psql -U "$DB_USER" -v ON_ERROR_STOP=1 -c "DROP DATABASE IF EXISTS $DB_NAME;"; then
  echo "FAILED: could not drop $DB_NAME"
  exit 1
fi
if ! docker exec "$CONTAINER" psql -U "$DB_USER" -v ON_ERROR_STOP=1 -c "CREATE DATABASE $DB_NAME;"; then
  echo "FAILED: could not create $DB_NAME"
  exit 1
fi

# Apply all migrations
echo "==> Applying migrations..."
if ! "$PROJECT_DIR/migrations/migrate.sh" "$CONTAINER" "$DB_NAME" "$DB_USER"; then
  echo ""
  echo "FAILED: migrations failed."
  exit 1
fi

# Seed test data
echo "==> Seeding test database..."
if ! (cd "$PROJECT_DIR" && DATABASE_URL="$DATABASE_URL" npx tsx prisma/seed.ts); then
  echo ""
  echo "FAILED: seed script failed."
  exit 1
fi

echo ""
echo "==> Test database ready."
echo "    - Migrations applied"
echo "    - Test data seeded"
echo "    - Database: $DB_NAME"
