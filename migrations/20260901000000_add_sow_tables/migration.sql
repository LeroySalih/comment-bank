-- SoW: Half Terms
CREATE TABLE IF NOT EXISTS "HalfTerm" (
    "id"           TEXT NOT NULL,
    "academicYear" TEXT NOT NULL,
    "label"        TEXT NOT NULL,  -- H1 .. H6
    "startDate"    DATE NOT NULL,
    "endDate"      DATE NOT NULL,
    CONSTRAINT "HalfTerm_pkey" PRIMARY KEY ("id")
);
CREATE UNIQUE INDEX IF NOT EXISTS "HalfTerm_year_label_key" ON "HalfTerm"("academicYear", "label");

-- SoW: Units that appear in the half-term grid
-- A unit is created manually by a teacher OR auto-created when a lesson is assigned.
-- When both sources share the same (classId, halfTermId, title) they are merged into one row.
CREATE TABLE IF NOT EXISTS "SowUnit" (
    "id"          TEXT NOT NULL,
    "classId"     TEXT NOT NULL,
    "halfTermId"  TEXT NOT NULL,
    "title"       TEXT NOT NULL,
    "comment"     TEXT,
    "isManual"    BOOLEAN NOT NULL DEFAULT false,
    CONSTRAINT "SowUnit_pkey" PRIMARY KEY ("id"),
    CONSTRAINT "SowUnit_class_fkey"    FOREIGN KEY ("classId")    REFERENCES "Class"("id")    ON DELETE CASCADE,
    CONSTRAINT "SowUnit_halfterm_fkey" FOREIGN KEY ("halfTermId") REFERENCES "HalfTerm"("id") ON DELETE CASCADE
);
CREATE UNIQUE INDEX IF NOT EXISTS "SowUnit_class_halfterm_title_key" ON "SowUnit"("classId", "halfTermId", "title");
CREATE INDEX IF NOT EXISTS "SowUnit_classId_idx"    ON "SowUnit"("classId");
CREATE INDEX IF NOT EXISTS "SowUnit_halfTermId_idx" ON "SowUnit"("halfTermId");

-- SoW: Lessons assigned to a class
CREATE TABLE IF NOT EXISTS "Lesson" (
    "id"                TEXT NOT NULL,
    "classId"           TEXT NOT NULL,
    "sowUnitId"         TEXT,
    "date"              DATE NOT NULL,
    "title"             TEXT NOT NULL,
    "score"             TEXT,
    "learningObjectives" TEXT,
    CONSTRAINT "Lesson_pkey"      PRIMARY KEY ("id"),
    CONSTRAINT "Lesson_class_fkey"   FOREIGN KEY ("classId")   REFERENCES "Class"("id")    ON DELETE CASCADE,
    CONSTRAINT "Lesson_sowUnit_fkey" FOREIGN KEY ("sowUnitId") REFERENCES "SowUnit"("id")  ON DELETE SET NULL
);
CREATE INDEX IF NOT EXISTS "Lesson_classId_idx"   ON "Lesson"("classId");
CREATE INDEX IF NOT EXISTS "Lesson_sowUnitId_idx" ON "Lesson"("sowUnitId");
