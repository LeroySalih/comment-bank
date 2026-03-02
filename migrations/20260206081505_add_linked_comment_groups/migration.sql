-- AlterTable
ALTER TABLE "Assignment" ADD COLUMN IF NOT EXISTS "linkedData" JSONB;

-- AlterTable
ALTER TABLE "CommentGroup" ADD COLUMN IF NOT EXISTS "isLinked" BOOLEAN NOT NULL DEFAULT false;
ALTER TABLE "CommentGroup" ADD COLUMN IF NOT EXISTS "linkedField" TEXT;

-- AlterTable
ALTER TABLE "CommonCommentGroup" ADD COLUMN IF NOT EXISTS "isLinked" BOOLEAN NOT NULL DEFAULT false;
ALTER TABLE "CommonCommentGroup" ADD COLUMN IF NOT EXISTS "linkedField" TEXT;
