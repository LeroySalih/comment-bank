-- Covers changes from 20260214132249 that may not have been applied

-- Drop paragraphPosition from CommonCommentGroup (added in 20260206011712, removed in 20260214132249)
ALTER TABLE "CommonCommentGroup" DROP COLUMN IF EXISTS "paragraphPosition";

-- Add commentFormat to Subject
ALTER TABLE "Subject" ADD COLUMN IF NOT EXISTS "commentFormat" TEXT;

-- Rename AppSetting key (safe to run repeatedly — only updates if old key exists)
UPDATE "AppSetting" SET "key" = 'comment_format_template' WHERE "key" = 'p2_wrapper_template';
