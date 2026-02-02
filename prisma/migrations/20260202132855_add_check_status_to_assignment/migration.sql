-- AlterTable
ALTER TABLE "Assignment" ADD COLUMN     "checkNote" TEXT,
ADD COLUMN     "checkStatus" TEXT NOT NULL DEFAULT 'not_required',
ADD COLUMN     "checkedAt" TIMESTAMP(3),
ADD COLUMN     "checkedById" TEXT;

-- CreateIndex
CREATE INDEX "Assignment_checkStatus_idx" ON "Assignment"("checkStatus");

-- AddForeignKey
ALTER TABLE "Assignment" ADD CONSTRAINT "Assignment_checkedById_fkey" FOREIGN KEY ("checkedById") REFERENCES "User"("id") ON DELETE SET NULL ON UPDATE CASCADE;
