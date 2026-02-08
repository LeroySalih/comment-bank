-- CreateTable
CREATE TABLE "CommonCommentGroup" (
    "id" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "title" TEXT NOT NULL DEFAULT '',
    "displayOrder" INTEGER NOT NULL DEFAULT 0,
    "paragraphPosition" TEXT NOT NULL,

    CONSTRAINT "CommonCommentGroup_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "CommonCommentOption" (
    "id" TEXT NOT NULL,
    "code" TEXT NOT NULL,
    "text" TEXT NOT NULL,
    "displayOrder" INTEGER NOT NULL DEFAULT 0,
    "groupId" TEXT NOT NULL,

    CONSTRAINT "CommonCommentOption_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "CommonPupilCode" (
    "id" TEXT NOT NULL,
    "assignmentId" TEXT NOT NULL,
    "commonGroupId" TEXT NOT NULL,
    "code" TEXT,

    CONSTRAINT "CommonPupilCode_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "AppSetting" (
    "key" TEXT NOT NULL,
    "value" TEXT NOT NULL,

    CONSTRAINT "AppSetting_pkey" PRIMARY KEY ("key")
);

-- CreateIndex
CREATE UNIQUE INDEX "CommonCommentGroup_name_key" ON "CommonCommentGroup"("name");

-- CreateIndex
CREATE INDEX "CommonCommentOption_groupId_idx" ON "CommonCommentOption"("groupId");

-- CreateIndex
CREATE UNIQUE INDEX "CommonCommentOption_groupId_code_key" ON "CommonCommentOption"("groupId", "code");

-- CreateIndex
CREATE INDEX "CommonPupilCode_assignmentId_idx" ON "CommonPupilCode"("assignmentId");

-- CreateIndex
CREATE INDEX "CommonPupilCode_commonGroupId_idx" ON "CommonPupilCode"("commonGroupId");

-- CreateIndex
CREATE UNIQUE INDEX "CommonPupilCode_assignmentId_commonGroupId_key" ON "CommonPupilCode"("assignmentId", "commonGroupId");

-- AddForeignKey
ALTER TABLE "CommonCommentOption" ADD CONSTRAINT "CommonCommentOption_groupId_fkey" FOREIGN KEY ("groupId") REFERENCES "CommonCommentGroup"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CommonPupilCode" ADD CONSTRAINT "CommonPupilCode_assignmentId_fkey" FOREIGN KEY ("assignmentId") REFERENCES "Assignment"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CommonPupilCode" ADD CONSTRAINT "CommonPupilCode_commonGroupId_fkey" FOREIGN KEY ("commonGroupId") REFERENCES "CommonCommentGroup"("id") ON DELETE CASCADE ON UPDATE CASCADE;
