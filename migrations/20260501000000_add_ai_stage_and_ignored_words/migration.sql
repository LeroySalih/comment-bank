-- Add aiStage to Assignment
ALTER TABLE "Assignment"
  ADD COLUMN IF NOT EXISTS "aiStage" TEXT DEFAULT NULL;
-- valid values: NULL | 'spag' | 'tone'

-- Create IgnoredWord table
CREATE TABLE IF NOT EXISTS "IgnoredWord" (
  id          TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
  "teacherId" TEXT NOT NULL REFERENCES "User"(id) ON DELETE CASCADE,
  word        TEXT NOT NULL,
  "createdAt" TIMESTAMP WITH TIME ZONE DEFAULT NOW(),
  UNIQUE ("teacherId", word)
);

CREATE INDEX IF NOT EXISTS "IgnoredWord_teacherId_idx" ON "IgnoredWord" ("teacherId");
