import { pool } from '@/lib/db';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';
import FormPupilRow from './_components/FormPupilRow';

export default async function FormPage({ params }: { params: Promise<{ formId: string }> }) {
  const { formId } = await params;
  const decodedFormId = decodeURIComponent(formId);
  const session = await getServerSession(authOptions);

  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userId = session?.user?.id;

  // Get classes the user is directly assigned to (for edit permissions)
  let userAssignedClassIds: string[] = [];

  if (!userIsAdmin && userId) {
    const { rows: assignedClassRows } = await pool.query(
      `SELECT c.id FROM "Class" c
       JOIN "_ClassToUser" cu ON cu."A" = c.id
       WHERE cu."B" = $1`,
      [userId]
    );
    userAssignedClassIds = assignedClassRows.map((r: any) => r.id);
  }

  // Fetch pupils in this form that the user has access to
  let pupilsQuery: string;
  let pupilsParams: unknown[];

  if (userIsAdmin) {
    pupilsQuery = `
      SELECT p.*
      FROM "Pupil" p
      WHERE p.form = $1 AND p."isActive" = true
        AND EXISTS (
          SELECT 1 FROM "Assignment" a WHERE a."pupilId" = p."admissionNumber"
        )
      ORDER BY p."lastName" ASC, p."firstName" ASC
    `;
    pupilsParams = [decodedFormId];
  } else if (userIsHoD) {
    pupilsQuery = `
      SELECT DISTINCT p.*
      FROM "Pupil" p
      JOIN "Assignment" a ON a."pupilId" = p."admissionNumber"
      JOIN "Class" c ON c.id = a."classId"
      JOIN "Subject" s ON s.id = c."subjectId"
      JOIN "_SubjectToUser" su ON su."A" = s.id
      WHERE p.form = $1 AND p."isActive" = true AND su."B" = $2
      ORDER BY p."lastName" ASC, p."firstName" ASC
    `;
    pupilsParams = [decodedFormId, userId];
  } else {
    pupilsQuery = `
      SELECT DISTINCT p.*
      FROM "Pupil" p
      JOIN "Assignment" a ON a."pupilId" = p."admissionNumber"
      JOIN "_ClassToUser" cu ON cu."A" = a."classId"
      WHERE p.form = $1 AND p."isActive" = true AND cu."B" = $2
      ORDER BY p."lastName" ASC, p."firstName" ASC
    `;
    pupilsParams = [decodedFormId, userId];
  }

  const { rows: pupilRows } = await pool.query(pupilsQuery, pupilsParams);

  if (pupilRows.length === 0) {
    return (
      <main className="flex-1 flex flex-col items-center justify-center min-h-screen bg-background-light dark:bg-background-dark">
        <div className="text-center">
          <h1 className="text-2xl font-bold text-slate-900 dark:text-white mb-2">Form Not Found</h1>
          <p className="text-slate-500 dark:text-slate-400 mb-4">No pupils found in form {decodedFormId} or you don&apos;t have access.</p>
          <Link href="/" className="text-primary hover:underline">Return to Dashboard</Link>
        </div>
      </main>
    );
  }

  const pupilIds = pupilRows.map((p: any) => p.admissionNumber);

  // Fetch all assignments for these pupils with class, subject, and comment groups
  const { rows: assignmentRows } = await pool.query(
    `SELECT a.*,
            c.id as class_id, c.name as class_name, c.year as class_year, c."subjectId" as class_subjectId,
            s.id as s_id, s.code as s_code, s.title as s_title, s."commentFormat" as s_commentFormat,
            s."studiedComment" as s_studiedComment
     FROM "Assignment" a
     JOIN "Class" c ON c.id = a."classId"
     JOIN "Subject" s ON s.id = c."subjectId"
     WHERE a."pupilId" = ANY($1::text[])
     ORDER BY s.code ASC`,
    [pupilIds]
  );

  const assignmentIds = assignmentRows.map((r: any) => r.id);

  // Fetch comment groups for all subjects involved
  const subjectIds = [...new Set(assignmentRows.map((r: any) => r.class_subjectId))] as string[];
  let commentGroupsBySubject = new Map<string, any[]>();

  if (subjectIds.length > 0) {
    const { rows: groupRows } = await pool.query(
      `SELECT * FROM "CommentGroup" WHERE "subjectId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
      [subjectIds]
    );

    if (groupRows.length > 0) {
      const groupIds = groupRows.map((g: any) => g.id);
      const { rows: optionRows } = await pool.query(
        `SELECT * FROM "CommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
        [groupIds]
      );
      const optionsByGroup = new Map<string, any[]>();
      for (const opt of optionRows) {
        const arr = optionsByGroup.get(opt.groupId) ?? [];
        arr.push(opt);
        optionsByGroup.set(opt.groupId, arr);
      }
      for (const g of groupRows) {
        (g as any).CommentOption = optionsByGroup.get(g.id) ?? [];
        const arr = commentGroupsBySubject.get(g.subjectId) ?? [];
        arr.push(g);
        commentGroupsBySubject.set(g.subjectId, arr);
      }
    }
  }

  // Fetch PupilCodes for all assignments
  let pupilCodesByAssignment = new Map<string, any[]>();

  if (assignmentIds.length > 0) {
    const { rows: pupilCodeRows } = await pool.query(
      `SELECT pc.*, cg.id as cg_id, cg.name as cg_name, cg."displayOrder" as cg_displayOrder,
              cg."subjectId" as cg_subjectId, cg.title as cg_title, cg."isLinked" as cg_isLinked,
              cg."linkedField" as cg_linkedField
       FROM "PupilCode" pc
       JOIN "CommentGroup" cg ON cg.id = pc."groupId"
       WHERE pc."assignmentId" = ANY($1::text[])
       ORDER BY cg."displayOrder" ASC`,
      [assignmentIds]
    );

    for (const row of pupilCodeRows) {
      const pc = {
        id: row.id,
        assignmentId: row.assignmentId,
        groupId: row.groupId,
        code: row.code,
        CommentGroup: {
          id: row.cg_id,
          name: row.cg_name,
          displayOrder: row.cg_displayOrder,
          subjectId: row.cg_subjectId,
          title: row.cg_title,
          isLinked: row.cg_isLinked,
          linkedField: row.cg_linkedField,
        }
      };
      const arr = pupilCodesByAssignment.get(row.assignmentId) ?? [];
      arr.push(pc);
      pupilCodesByAssignment.set(row.assignmentId, arr);
    }
  }

  // Group assignments by pupilId
  const assignmentsByPupil = new Map<string, any[]>();
  for (const row of assignmentRows) {
    const assignment = {
      id: row.id,
      pupilId: row.pupilId,
      classId: row.classId,
      eoyLevel: row.eoyLevel,
      targetLevel: row.targetLevel,
      actualLevel: row.actualLevel,
      finalComment: row.finalComment,
      linkedData: row.linkedData,
      checkStatus: row.checkStatus,
      checkNote: row.checkNote,
      checkedAt: row.checkedAt,
      checkedById: row.checkedById,
      PupilCode: pupilCodesByAssignment.get(row.id) ?? [],
      Class: {
        id: row.class_id,
        name: row.class_name,
        year: row.class_year,
        subjectId: row.class_subjectId,
        Subject: {
          id: row.s_id,
          code: row.s_code,
          title: row.s_title,
          commentFormat: row.s_commentFormat,
          studiedComment: row.s_studiedComment,
          CommentGroup: commentGroupsBySubject.get(row.class_subjectId) ?? [],
        }
      }
    };
    const arr = assignmentsByPupil.get(row.pupilId) ?? [];
    arr.push(assignment);
    assignmentsByPupil.set(row.pupilId, arr);
  }

  // Decrypt pupil names and build full pupil objects
  const decryptedPupils = pupilRows.map((pupil: any) => ({
    admissionNumber: pupil.admissionNumber,
    firstName: decrypt(pupil.firstName),
    lastName: decrypt(pupil.lastName),
    gender: pupil.gender,
    isActive: pupil.isActive,
    form: pupil.form,
    Assignment: assignmentsByPupil.get(pupil.admissionNumber) ?? [],
  }));

  decryptedPupils.sort((a: any, b: any) =>
    a.lastName.localeCompare(b.lastName) || a.firstName.localeCompare(b.firstName)
  );

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-slate-50 dark:bg-slate-950">
      {/* Page Heading */}
      <div className="flex flex-wrap items-center justify-between gap-4 p-6 bg-white dark:bg-[#1a222c] border-b border-slate-200 dark:border-[#2d3748]">
        <div className="flex flex-col gap-1">
          <div className="flex items-center gap-2">
            <Link href="/" className="text-gray-500 hover:text-gray-900 dark:hover:text-white transition-colors">
              <ArrowLeft className="w-5 h-5" />
            </Link>
            <h1 className="text-[#111418] dark:text-white text-3xl font-black leading-tight tracking-[-0.033em]">
              Form: {decodedFormId}
            </h1>
          </div>
          <p className="text-[#617289] dark:text-gray-400 text-sm font-normal leading-normal">
            {decryptedPupils.length} Pupils
          </p>
        </div>
      </div>

      {/* Pupils Grid */}
      <div className="p-6">
        <div className="bg-white dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-800 overflow-hidden">
          {/* Header Row */}
          <div className="grid grid-cols-[200px_1fr_1fr] gap-4 px-4 py-3 bg-slate-100 dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 text-xs font-semibold text-slate-600 dark:text-slate-400 uppercase tracking-wider">
            <div>Pupil</div>
            <div>Subjects & Codes</div>
            <div>Comments</div>
          </div>

          {/* Pupil Rows */}
          <div className="divide-y divide-slate-100 dark:divide-slate-800">
            {decryptedPupils.map((pupil: any) => {
              // For admin users, get all class IDs from the pupil's assignments
              const editableClassIds = userIsAdmin
                ? pupil.Assignment.map((a: any) => a.Class.id)
                : userAssignedClassIds;

              return (
                <FormPupilRow
                  key={pupil.admissionNumber}
                  pupil={pupil}
                  userAssignedClassIds={editableClassIds}
                />
              );
            })}
          </div>
        </div>
      </div>
    </main>
  );
}
