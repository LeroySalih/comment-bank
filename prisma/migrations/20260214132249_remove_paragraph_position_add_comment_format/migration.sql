/*
  Warnings:

  - You are about to drop the column `paragraphPosition` on the `CommonCommentGroup` table. All the data in the column will be lost.

*/
-- AlterTable
ALTER TABLE "CommonCommentGroup" DROP COLUMN "paragraphPosition";

-- AlterTable
ALTER TABLE "Subject" ADD COLUMN     "commentFormat" TEXT;

-- Rename AppSetting key from p2_wrapper_template to comment_format_template
UPDATE "AppSetting" SET "key" = 'comment_format_template' WHERE "key" = 'p2_wrapper_template';
