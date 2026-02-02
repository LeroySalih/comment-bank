-- AlterTable
ALTER TABLE "Class" ADD COLUMN "year" TEXT;

-- AlterTable
ALTER TABLE "Course" ADD COLUMN "subject" TEXT;

-- AlterTable
ALTER TABLE "Student" ADD COLUMN "actualLevel" TEXT;
ALTER TABLE "Student" ADD COLUMN "eoyLevel" TEXT;
ALTER TABLE "Student" ADD COLUMN "targetLevel" TEXT;

-- CreateTable
CREATE TABLE "_ClassToUser" (
    "A" TEXT NOT NULL,
    "B" TEXT NOT NULL,
    CONSTRAINT "_ClassToUser_A_fkey" FOREIGN KEY ("A") REFERENCES "Class" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "_ClassToUser_B_fkey" FOREIGN KEY ("B") REFERENCES "User" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- RedefineTables
PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;
CREATE TABLE "new_CommentGroup" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "name" TEXT NOT NULL,
    "order" INTEGER NOT NULL DEFAULT 0,
    "courseId" TEXT NOT NULL,
    CONSTRAINT "CommentGroup_courseId_fkey" FOREIGN KEY ("courseId") REFERENCES "Course" ("id") ON DELETE RESTRICT ON UPDATE CASCADE
);
INSERT INTO "new_CommentGroup" ("courseId", "id", "name") SELECT "courseId", "id", "name" FROM "CommentGroup";
DROP TABLE "CommentGroup";
ALTER TABLE "new_CommentGroup" RENAME TO "CommentGroup";
CREATE UNIQUE INDEX "CommentGroup_courseId_name_key" ON "CommentGroup"("courseId", "name");
PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;

-- CreateIndex
CREATE UNIQUE INDEX "_ClassToUser_AB_unique" ON "_ClassToUser"("A", "B");

-- CreateIndex
CREATE INDEX "_ClassToUser_B_index" ON "_ClassToUser"("B");
