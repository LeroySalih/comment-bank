/**
 * Shared DB type interfaces replacing Prisma-generated types.
 *
 * Naming convention: Db<ModelName> (scalar fields only).
 * Composite "with relations" interfaces are defined at the bottom.
 *
 * Type mapping from Prisma/PostgreSQL via `pg`:
 *   String    -> string
 *   Boolean   -> boolean
 *   Int       -> number
 *   DateTime  -> Date
 *   Json      -> Record<string, any> | null
 *   Optional  -> T | null
 */

// ---------------------------------------------------------------------------
// Scalar interfaces (one per Prisma model)
// ---------------------------------------------------------------------------

export interface DbAssignment {
  id: string;
  pupilId: string;
  classId: string;
  eoyLevel: string | null;
  targetLevel: string | null;
  actualLevel: string | null;
  finalComment: string | null;
  linkedData: Record<string, any> | null;
  checkStatus: string;
  checkNote: string | null;
  checkedAt: Date | null;
  checkedById: string | null;
}

export interface DbClass {
  id: string;
  name: string;
  year: string | null;
  subjectId: string;
}

export interface DbCommentGroup {
  id: string;
  name: string;
  displayOrder: number;
  subjectId: string;
  title: string;
  isLinked: boolean;
  linkedField: string | null;
}

export interface DbCommentOption {
  id: string;
  code: string;
  text: string;
  displayOrder: number;
  groupId: string;
}

export interface DbPupil {
  admissionNumber: string;
  firstName: string;
  lastName: string;
  gender: string;
  isActive: boolean;
  form: string | null;
}

export interface DbPupilCode {
  id: string;
  assignmentId: string;
  groupId: string;
  code: string | null;
}

export interface DbRole {
  id: string;
  name: string;
}

export interface DbSubject {
  id: string;
  code: string;
  title: string | null;
  studiedComment: string | null;
  commentFormat: string | null;
}

export interface DbUser {
  id: string;
  username: string;
  password: string;
  isActive: boolean;
}

export interface DbDeadline {
  id: string;
  title: string;
  date: Date;
  description: string | null;
  isActive: boolean;
  createdAt: Date;
}

export interface DbCommonCommentGroup {
  id: string;
  name: string;
  title: string;
  displayOrder: number;
  isLinked: boolean;
  linkedField: string | null;
}

export interface DbCommonCommentOption {
  id: string;
  code: string;
  text: string;
  displayOrder: number;
  groupId: string;
}

export interface DbCommonPupilCode {
  id: string;
  assignmentId: string;
  commonGroupId: string;
  code: string | null;
}

export interface DbAppSetting {
  key: string;
  value: string;
}

export interface DbAuditLog {
  id: string;
  userId: string | null;
  username: string | null;
  action: string;
  entityType: string | null;
  entityId: string | null;
  details: string | null;
  ipAddress: string | null;
  userAgent: string | null;
  createdAt: Date;
}

// ---------------------------------------------------------------------------
// Composite "with relations" interfaces
// ---------------------------------------------------------------------------

export interface DbUserWithRoles extends DbUser {
  Role: DbRole[];
}
