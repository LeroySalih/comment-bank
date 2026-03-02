-- AlterTable
ALTER TABLE "Assignment" ADD COLUMN IF NOT EXISTS "checkNote" TEXT;
ALTER TABLE "Assignment" ADD COLUMN IF NOT EXISTS "checkStatus" TEXT NOT NULL DEFAULT 'not_required';
ALTER TABLE "Assignment" ADD COLUMN IF NOT EXISTS "checkedAt" TIMESTAMP(3);
ALTER TABLE "Assignment" ADD COLUMN IF NOT EXISTS "checkedById" TEXT;

-- CreateIndex
CREATE INDEX IF NOT EXISTS "Assignment_checkStatus_idx" ON "Assignment"("checkStatus");

-- AddForeignKey
DO $$ BEGIN
  ALTER TABLE "Assignment" ADD CONSTRAINT "Assignment_checkedById_fkey" FOREIGN KEY ("checkedById") REFERENCES "User"("id") ON DELETE SET NULL ON UPDATE CASCADE;
EXCEPTION WHEN duplicate_object THEN NULL;
END $$;
