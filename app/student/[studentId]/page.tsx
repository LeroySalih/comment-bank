import { pool } from '@/lib/db';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import CommentEditor from '@/components/CommentEditor';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';
import { unstable_noStore as noStore } from 'next/cache';

// Disable caching to always get fresh data
export const dynamic = 'force-dynamic';
export const revalidate = 0;

export default async function StudentPage({ params }: { params: Promise<{ studentId: string }> }) {
  noStore(); // Disable all caching for this request

  const { studentId: assignmentId } = await params;
  const session = await getServerSession(authOptions);

  // Fetch assignment with class, subject, and pupil
  const { rows: assignmentRows } = await pool.query(
    `SELECT a.*,
            p."admissionNumber" as "pupil_admissionNumber",
            p."firstName" as "pupil_firstName",
            p."lastName" as "pupil_lastName",
            p.gender as "pupil_gender",
            p."isActive" as "pupil_isActive",
            p.form as "pupil_form",
            c.id as "class_id", c.name as "class_name", c.year as "class_year", c."subjectId" as "class_subjectId",
            s.id as "s_id", s.code as "s_code", s.title as "s_title", s."commentFormat" as "s_commentFormat",
            s."studiedComment" as "s_studiedComment"
     FROM "Assignment" a
     JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
     JOIN "Class" c ON c.id = a."classId"
     JOIN "Subject" s ON s.id = c."subjectId"
     WHERE a.id = $1`,
    [assignmentId]
  );

  if (assignmentRows.length === 0) {
    return <div>Assignment not found</div>;
  }

  const row = assignmentRows[0];

  // Fetch teachers for authorization check
  const { rows: teacherRows } = await pool.query(
    `SELECT u.id FROM "User" u
     JOIN "_ClassToUser" cu ON cu."B" = u.id
     WHERE cu."A" = $1`,
    [row.class_id]
  );

  // Authorization check
  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userIsTeacher = isTeacher(session?.user);

  // HoD can review comments (admins also have this capability)
  const canReviewComments = userIsHoD || userIsAdmin;

  if (userIsTeacher && !userIsAdmin && !userIsHoD) {
    const isAssigned = teacherRows.some((t: any) => t.id === session?.user?.id);
    if (!isAssigned) {
      return <div>Student not found or access denied</div>;
    }
  }

  // Fetch comment groups for subject with options
  const { rows: groupRows } = await pool.query(
    `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 ORDER BY "displayOrder" ASC`,
    [row.class_subjectId]
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
    }
  }

  // Fetch PupilCodes for this assignment
  const { rows: pupilCodeRows } = await pool.query(
    `SELECT * FROM "PupilCode" WHERE "assignmentId" = $1`,
    [assignmentId]
  );

  // Fetch CommonPupilCodes for this assignment
  const { rows: commonPupilCodeRows } = await pool.query(
    `SELECT * FROM "CommonPupilCode" WHERE "assignmentId" = $1`,
    [assignmentId]
  );

  const subject = {
    id: row.s_id,
    code: row.s_code,
    title: row.s_title,
    commentFormat: row.s_commentFormat,
    studiedComment: row.s_studiedComment,
    CommentGroup: groupRows,
  };

  const groups = groupRows;

  // Fetch Common Comment Groups
  const { rows: commonGroupRows } = await pool.query(
    `SELECT * FROM "CommonCommentGroup" ORDER BY "displayOrder" ASC`
  );

  if (commonGroupRows.length > 0) {
    const cgIds = commonGroupRows.map((g: any) => g.id);
    const { rows: cgOptionRows } = await pool.query(
      `SELECT * FROM "CommonCommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
      [cgIds]
    );
    const optionsByCg = new Map<string, any[]>();
    for (const opt of cgOptionRows) {
      const arr = optionsByCg.get(opt.groupId) ?? [];
      arr.push(opt);
      optionsByCg.set(opt.groupId, arr);
    }
    for (const g of commonGroupRows) {
      (g as any).CommonCommentOption = optionsByCg.get(g.id) ?? [];
    }
  }

  const commonGroups = commonGroupRows;

  // Fetch format template
  const { rows: settingRows } = await pool.query(
    `SELECT value FROM "AppSetting" WHERE key = 'comment_format_template'`
  );
  const formatTemplate = settingRows[0]?.value || '';
  const subjectFormat = row.s_commentFormat || null;

  // Decrypt pupil names
  const pupil = {
    admissionNumber: row.pupil_admissionNumber,
    firstName: decrypt(row.pupil_firstName),
    lastName: decrypt(row.pupil_lastName),
    gender: row.pupil_gender,
    isActive: row.pupil_isActive,
    form: row.pupil_form,
  };

  // Build decrypted assignment shape expected by CommentEditor
  const decryptedAssignment = {
    id: row.id,
    pupilId: row.pupilId,
    classId: row.classId,
    eoyLevel: row.eoyLevel,
    targetLevel: row.targetLevel,
    actualLevel: row.actualLevel,
    finalComment: row.finalComment ? decrypt(row.finalComment) : null,
    linkedData: row.linkedData,
    checkStatus: row.checkStatus,
    checkNote: row.checkNote,
    checkedAt: row.checkedAt,
    checkedById: row.checkedById,
    aiStage: row.aiStage ?? null,
    Pupil: pupil,
    PupilCode: pupilCodeRows,
    CommonPupilCode: commonPupilCodeRows,
    Class: {
      id: row.class_id,
      name: row.class_name,
      year: row.class_year,
      subjectId: row.class_subjectId,
      User: teacherRows,
      Subject: subject,
    },
  };

  return (
    <main className="min-h-screen flex flex-col bg-background-light dark:bg-background-dark">
      {/* Breadcrumbs */}
      <div className="px-10 py-4 flex items-center justify-between">
        <div className="flex flex-wrap gap-2 items-center">
            <Link href="/" className="text-[#617289] dark:text-gray-400 text-sm font-medium hover:text-primary transition-colors">Dashboard</Link>
            <span className="text-[#617289] dark:text-gray-600 material-symbols-outlined text-sm">chevron_right</span>
            <Link href={`/class/${row.class_id}`} className="text-[#617289] dark:text-gray-400 text-sm font-medium hover:text-primary transition-colors">{row.class_name}</Link>
            <span className="text-[#617289] dark:text-gray-600 material-symbols-outlined text-sm">chevron_right</span>
            <span className="text-[#111418] dark:text-white text-sm font-semibold">{pupil.firstName} {pupil.lastName}</span>
        </div>
        <div className="flex items-center gap-2 text-xs text-[#617289] dark:text-gray-400 italic">
            <span className="material-symbols-outlined text-xs">sync</span>
            Auto-save enabled
        </div>
      </div>

      {/* Page Header */}
      <div className="px-10 pb-6 flex flex-wrap justify-between items-end gap-3">
        <div className="flex min-w-72 flex-col gap-1">
            <div className="flex items-center gap-3">
                <h1 className="text-[#111418] dark:text-white text-4xl font-black leading-tight tracking-[-0.033em]">Edit Report: {pupil.firstName} {pupil.lastName}</h1>
            </div>
            <p className="text-[#617289] dark:text-gray-400 text-base font-normal">{subject.title} • {pupil.gender}</p>
        </div>
      </div>

      <div className="flex-1 px-10 pb-10 flex gap-8">
        <CommentEditor
          assignment={decryptedAssignment}
          subject={subject}
          groups={groups}
          isHoD={canReviewComments}
          commonGroups={commonGroups}
          formatTemplate={formatTemplate}
          subjectFormat={subjectFormat}
        />
      </div>
    </main>
  );
}
