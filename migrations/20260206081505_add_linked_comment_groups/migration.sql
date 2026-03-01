-- AlterTable
ALTER TABLE "Assignment" ADD COLUMN     "linkedData" JSONB;

-- AlterTable
ALTER TABLE "CommentGroup" ADD COLUMN     "isLinked" BOOLEAN NOT NULL DEFAULT false,
ADD COLUMN     "linkedField" TEXT;

-- AlterTable
ALTER TABLE "CommonCommentGroup" ADD COLUMN     "isLinked" BOOLEAN NOT NULL DEFAULT false,
ADD COLUMN     "linkedField" TEXT;
