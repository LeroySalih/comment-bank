/*
  Warnings:

  - You are about to drop the `Course` table. If the table is not empty, all the data it contains will be lost.
  - You are about to drop the `Student` table. If the table is not empty, all the data it contains will be lost.
  - You are about to drop the column `courseId` on the `Class` table. All the data in the column will be lost.
  - You are about to drop the column `courseId` on the `CommentGroup` table. All the data in the column will be lost.
  - Added the required column `subjectId` to the `Class` table without a default value. This is not possible if the table is not empty.
  - Added the required column `subjectId` to the `CommentGroup` table without a default value. This is not possible if the table is not empty.

*/
-- DropIndex
DROP INDEX "Course_name_key";

-- DropTable
PRAGMA foreign_keys=off;
DROP TABLE "Course";
PRAGMA foreign_keys=on;

-- DropTable
PRAGMA foreign_keys=off;
DROP TABLE "Student";
PRAGMA foreign_keys=on;

-- CreateTable
CREATE TABLE "Subject" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "name" TEXT NOT NULL,
    "subject" TEXT,
    "studiedComment" TEXT,
    "ownerId" TEXT,
    CONSTRAINT "Subject_ownerId_fkey" FOREIGN KEY ("ownerId") REFERENCES "User" ("id") ON DELETE SET NULL ON UPDATE CASCADE
);

-- CreateTable
CREATE TABLE "Pupil" (
    "admissionNumber" TEXT NOT NULL PRIMARY KEY,
    "firstName" TEXT NOT NULL,
    "lastName" TEXT NOT NULL,
    "gender" TEXT NOT NULL,
    "isActive" BOOLEAN NOT NULL DEFAULT true
);

-- CreateTable
CREATE TABLE "Assignment" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "pupilId" TEXT NOT NULL,
    "classId" TEXT NOT NULL,
    "eoyLevel" TEXT,
    "targetLevel" TEXT,
    "actualLevel" TEXT,
    "finalComment" TEXT,
    CONSTRAINT "Assignment_pupilId_fkey" FOREIGN KEY ("pupilId") REFERENCES "Pupil" ("admissionNumber") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "Assignment_classId_fkey" FOREIGN KEY ("classId") REFERENCES "Class" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- CreateTable
CREATE TABLE "PupilCode" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "assignmentId" TEXT NOT NULL,
    "groupId" TEXT NOT NULL,
    "code" TEXT,
    CONSTRAINT "PupilCode_assignmentId_fkey" FOREIGN KEY ("assignmentId") REFERENCES "Assignment" ("id") ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT "PupilCode_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "CommentGroup" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);

-- RedefineTables
PRAGMA defer_foreign_keys=ON;
PRAGMA foreign_keys=OFF;
CREATE TABLE "new_Class" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "name" TEXT NOT NULL,
    "year" TEXT,
    "subjectId" TEXT NOT NULL,
    CONSTRAINT "Class_subjectId_fkey" FOREIGN KEY ("subjectId") REFERENCES "Subject" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);
INSERT INTO "new_Class" ("id", "name", "year") SELECT "id", "name", "year" FROM "Class";
DROP TABLE "Class";
ALTER TABLE "new_Class" RENAME TO "Class";
CREATE UNIQUE INDEX "Class_name_key" ON "Class"("name");
CREATE TABLE "new_CommentGroup" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "name" TEXT NOT NULL,
    "displayOrder" INTEGER NOT NULL DEFAULT 0,
    "subjectId" TEXT NOT NULL,
    CONSTRAINT "CommentGroup_subjectId_fkey" FOREIGN KEY ("subjectId") REFERENCES "Subject" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);
INSERT INTO "new_CommentGroup" ("displayOrder", "id", "name") SELECT "displayOrder", "id", "name" FROM "CommentGroup";
DROP TABLE "CommentGroup";
ALTER TABLE "new_CommentGroup" RENAME TO "CommentGroup";
CREATE UNIQUE INDEX "CommentGroup_subjectId_name_key" ON "CommentGroup"("subjectId", "name");
CREATE TABLE "new_CommentOption" (
    "id" TEXT NOT NULL PRIMARY KEY,
    "code" TEXT NOT NULL,
    "text" TEXT NOT NULL,
    "displayOrder" INTEGER NOT NULL DEFAULT 0,
    "groupId" TEXT NOT NULL,
    CONSTRAINT "CommentOption_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "CommentGroup" ("id") ON DELETE CASCADE ON UPDATE CASCADE
);
INSERT INTO "new_CommentOption" ("code", "displayOrder", "groupId", "id", "text") SELECT "code", "displayOrder", "groupId", "id", "text" FROM "CommentOption";
DROP TABLE "CommentOption";
ALTER TABLE "new_CommentOption" RENAME TO "CommentOption";
CREATE UNIQUE INDEX "CommentOption_groupId_code_key" ON "CommentOption"("groupId", "code");
PRAGMA foreign_keys=ON;
PRAGMA defer_foreign_keys=OFF;

-- CreateIndex
CREATE UNIQUE INDEX "Subject_name_key" ON "Subject"("name");

-- CreateIndex
CREATE UNIQUE INDEX "PupilCode_assignmentId_groupId_key" ON "PupilCode"("assignmentId", "groupId");
