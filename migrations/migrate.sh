#!/usr/bin/env bash
# migrate.sh — Apply pending migrations to a Dockerised PostgreSQL instance
#
# Usage:
#   ./migrations/migrate.sh [CONTAINER] [DB_NAME] [DB_USER]
#
# Defaults:
#   CONTAINER = postgres17
#   DB_NAME   = comment_bank
#   DB_USER   = postgres
#
# The script compares migration directories in this folder against the
# _prisma_migrations table and applies any that are missing, in name order.

set -euo pipefail

CONTAINER="${1:-postgres17}"
DB_NAME="${2:-comment_bank}"
DB_USER="${3:-postgres}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# ── helpers ──────────────────────────────────────────────────────────────────

psql_exec() {
  docker exec -i "$CONTAINER" psql -U "$DB_USER" -d "$DB_NAME" "$@"
}

generate_uuid() {
  python3 -c "import uuid; print(uuid.uuid4())"
}

now_utc() {
  python3 -c "
from datetime import datetime, timezone
print(datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S.%f+00'))
"
}

checksum_of() {
  local file="$1"
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum "$file" | awk '{print $1}'
  else
    shasum -a 256 "$file" | awk '{print $1}'
  fi
}

# ── pre-flight ────────────────────────────────────────────────────────────────

echo "=== Migration Runner ==="
echo "  Container : $CONTAINER"
echo "  Database  : $DB_NAME"
echo "  User      : $DB_USER"
echo ""

# Verify the container is running
if ! docker inspect -f '{{.State.Running}}' "$CONTAINER" 2>/dev/null | grep -q true; then
  echo "ERROR: Container '$CONTAINER' is not running."
  exit 1
fi

# Create the tracking table if it doesn't exist yet
psql_exec -q -c "
  CREATE TABLE IF NOT EXISTS \"_prisma_migrations\" (
    id                  VARCHAR(36)  NOT NULL,
    checksum            VARCHAR(64)  NOT NULL,
    finished_at         TIMESTAMPTZ,
    migration_name      VARCHAR(255) NOT NULL,
    logs                TEXT,
    rolled_back_at      TIMESTAMPTZ,
    started_at          TIMESTAMPTZ  NOT NULL DEFAULT now(),
    applied_steps_count INTEGER      NOT NULL DEFAULT 0,
    PRIMARY KEY (id)
  );
"

# ── discover migrations ───────────────────────────────────────────────────────

# Get the set of already-applied migration names
APPLIED=$(psql_exec -t -A -c \
  "SELECT migration_name FROM \"_prisma_migrations\" ORDER BY migration_name;" \
  2>/dev/null || true)

# Collect pending migrations (directories containing migration.sql), sorted
PENDING=()
while IFS= read -r -d '' dir; do
  name="$(basename "$dir")"
  # skip non-migration entries (e.g. migration_lock.toml)
  [[ "$name" == *.* ]] && continue
  [[ -f "$dir/migration.sql" ]] || continue
  if ! grep -qx "$name" <<< "$APPLIED"; then
    PENDING+=("$name")
  fi
done < <(find "$SCRIPT_DIR" -mindepth 1 -maxdepth 1 -type d -print0 | sort -z)

if [[ ${#PENDING[@]} -eq 0 ]]; then
  echo "✓ All migrations are already applied."
  exit 0
fi

echo "${#PENDING[@]} pending migration(s):"
for name in "${PENDING[@]}"; do
  echo "  • $name"
done
echo ""

# ── apply migrations ──────────────────────────────────────────────────────────

APPLIED_COUNT=0

for name in "${PENDING[@]}"; do
  sql_file="$SCRIPT_DIR/$name/migration.sql"
  tmp_path="/tmp/_migration_${name}.sql"

  printf "Applying %-60s " "$name ..."

  checksum="$(checksum_of "$sql_file")"
  id="$(generate_uuid)"
  started_at="$(now_utc)"

  # Copy SQL into the container and execute it
  docker cp "$sql_file" "$CONTAINER:$tmp_path" >/dev/null

  if docker exec "$CONTAINER" psql -U "$DB_USER" -d "$DB_NAME" \
       -v ON_ERROR_STOP=1 -q -f "$tmp_path" 2>/tmp/migration_err.txt; then

    finished_at="$(now_utc)"

    # Record the migration in the tracking table
    psql_exec -q -c "
      INSERT INTO \"_prisma_migrations\"
        (id, checksum, finished_at, migration_name, logs, started_at, applied_steps_count)
      VALUES
        ('$id', '$checksum', '$finished_at', '$name', NULL, '$started_at', 1);
    "

    echo "✓"
    APPLIED_COUNT=$((APPLIED_COUNT + 1))

  else
    echo "✗ FAILED"
    echo ""
    echo "--- Error output ---"
    cat /tmp/migration_err.txt
    echo "--- End error output ---"
    echo ""
    echo "Migration '$name' failed. Stopping."
    exit 1
  fi

  # Clean up temp file in container
  docker exec "$CONTAINER" rm -f "$tmp_path" >/dev/null 2>&1 || true
done

# ── summary ───────────────────────────────────────────────────────────────────

echo ""
echo "=== Done — $APPLIED_COUNT migration(s) applied ==="
