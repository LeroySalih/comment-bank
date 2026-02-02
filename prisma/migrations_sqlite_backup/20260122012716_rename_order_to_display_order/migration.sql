/*
  Warnings:

  - You are about to drop the column `order` on the `CommentGroup` table. All the data in the column will be lost.
  - You are about to drop the column `order` on the `CommentOption` table. All the data in the column will be lost.

*/
-- RedefineTables
PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;
CREATE TABLE "new_CommentGroup" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "name" TEXT NOT NULL,
    "displayOrder" INTEGER NOT NULL DEFAULT 0,
    "courseId" TEXT NOT NULL,
    CONSTRAINT "CommentGroup_courseId_fkey" FOREIGN KEY ("courseId") REFERENCES "Course" ("id") ON DELETE RESTRICT ON UPDATE CASCADE
);
INSERT INTO "new_CommentGroup" ("courseId", "id", "name") SELECT "courseId", "id", "name" FROM "CommentGroup";
DROP TABLE "CommentGroup";
ALTER TABLE "new_CommentGroup" RENAME TO "CommentGroup";
CREATE UNIQUE INDEX "CommentGroup_courseId_name_key" ON "CommentGroup"("courseId", "name");
CREATE TABLE "new_CommentOption" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "code" TEXT NOT NULL,
    "text" TEXT NOT NULL,
    "displayOrder" INTEGER NOT NULL DEFAULT 0,
    "groupId" TEXT NOT NULL,
    CONSTRAINT "CommentOption_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "CommentGroup" ("id") ON DELETE RESTRICT ON UPDATE CASCADE
);
INSERT INTO "new_CommentOption" ("code", "groupId", "id", "text") SELECT "code", "groupId", "id", "text" FROM "CommentOption";
DROP TABLE "CommentOption";
ALTER TABLE "new_CommentOption" RENAME TO "CommentOption";
CREATE UNIQUE INDEX "CommentOption_groupId_code_key" ON "CommentOption"("groupId", "code");
PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;
