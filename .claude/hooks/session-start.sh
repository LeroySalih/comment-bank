#!/bin/bash
set -euo pipefail

# Only run in remote (Claude Code on the web) environments
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

WORKTREE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
MAIN_REPO_DIR="$(cd "$WORKTREE_DIR/../.." && pwd)"

echo "==> [SoW worktree] Session start hook"
echo "    Worktree : $WORKTREE_DIR"
echo "    Main repo: $MAIN_REPO_DIR"

# ── 1. Symlink node_modules from main repo if available, else install ─────────
if [ -d "$MAIN_REPO_DIR/node_modules" ]; then
  echo "==> Symlinking node_modules from main repo..."
  ln -sfn "$MAIN_REPO_DIR/node_modules" "$WORKTREE_DIR/node_modules"
else
  echo "==> Installing npm dependencies..."
  cd "$WORKTREE_DIR"
  npm install
fi

# ── 2. Copy .env.local from main repo if not already present ─────────────────
ENV_TARGET="$WORKTREE_DIR/.env.local"
if [ ! -f "$ENV_TARGET" ]; then
  for candidate in \
    "$MAIN_REPO_DIR/.env.local" \
    "$MAIN_REPO_DIR/.env" \
    "$MAIN_REPO_DIR/.env.development.local"; do
    if [ -f "$candidate" ]; then
      echo "==> Copying env from $candidate"
      cp "$candidate" "$ENV_TARGET"
      break
    fi
  done

  if [ ! -f "$ENV_TARGET" ]; then
    echo "⚠️  No .env file found in main repo. Create $ENV_TARGET with:"
    echo "    DATABASE_URL=postgresql://..."
    echo "    NEXTAUTH_SECRET=..."
    echo "    NEXTAUTH_URL=http://localhost:3000"
  fi
fi

# ── 3. Apply pending migrations if DATABASE_URL is available ──────────────────
# Load env so DATABASE_URL is accessible
if [ -f "$ENV_TARGET" ]; then
  set -a
  source "$ENV_TARGET"
  set +a
fi

if [ -n "${DATABASE_URL:-}" ]; then
  echo "==> Applying pending SQL migrations..."
  cd "$WORKTREE_DIR"
  for dir in migrations/*/; do
    sql="$dir/migration.sql"
    [ -f "$sql" ] || continue
    name="$(basename "$dir")"
    # Check if already applied via _prisma_migrations tracking table
    applied=$(psql "$DATABASE_URL" -At -c \
      "SELECT 1 FROM \"_prisma_migrations\" WHERE migration_name = '$name' LIMIT 1" 2>/dev/null || echo "")
    if [ -z "$applied" ]; then
      echo "  Applying: $name"
      psql "$DATABASE_URL" -v ON_ERROR_STOP=1 -f "$sql"
      psql "$DATABASE_URL" -c \
        "INSERT INTO \"_prisma_migrations\" (id, checksum, finished_at, migration_name, started_at, applied_steps_count)
         VALUES (gen_random_uuid()::text, 'manual', now(), '$name', now(), 1)
         ON CONFLICT DO NOTHING" 2>/dev/null || true
    fi
  done
  echo "==> Migrations done."
else
  echo "⚠️  DATABASE_URL not set — skipping migrations."
fi

echo "==> Session start hook complete."
