import { pool } from '@/lib/db';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import Tooltip from '@/components/Tooltip';
import { getServerSession } from 'next-auth';
import { authOptions } from '../../api/auth/[...nextauth]/route';
import { isAdmin, isHoD, isTeacher } from '@/lib/access-control';
import { decrypt } from '@/lib/encryption';
import StudentMatrixRow from './_components/StudentMatrixRow';

export default async function ClassPage({ params }: { params: Promise<{ classId: string }> }) {
  const { classId } = await params;
  const session = await getServerSession(authOptions);

  // Fetch class with subject and teachers
  const { rows: classRows } = await pool.query(
    `SELECT c.*,
            s.id as s_id, s.code as s_code, s.title as s_title,
            s."commentFormat" as "s_commentFormat"
     FROM "Class" c
     JOIN "Subject" s ON s.id = c."subjectId"
     WHERE c.id = $1`,
    [classId]
  );

  if (classRows.length === 0) {
    return <div>Class not found</div>;
  }

  const classRow = classRows[0];

  // Fetch teachers for this class
  const { rows: teacherRows } = await pool.query(
    `SELECT u.id FROM "User" u
     JOIN "_ClassToUser" cu ON cu."B" = u.id
     WHERE cu."A" = $1`,
    [classId]
  );

  const cls = {
    id: classRow.id,
    name: classRow.name,
    year: classRow.year,
    subjectId: classRow.subjectId,
    User: teacherRows,
    Subject: {
      id: classRow.s_id,
      code: classRow.s_code,
      title: classRow.s_title,
      commentFormat: classRow.s_commentFormat,
    }
  };

  // Authorization check
  const userIsAdmin = isAdmin(session?.user);
  const userIsHoD = isHoD(session?.user);
  const userIsTeacher = isTeacher(session?.user);

  if (userIsTeacher && !userIsAdmin && !userIsHoD) {
    const isAssigned = cls.User.some((t: any) => t.id === session?.user?.id);
    if (!isAssigned) {
      return <div>Class not found or access denied</div>;
    }
  }

  // Fetch comment groups for subject
  const { rows: groupRows } = await pool.query(
    `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 ORDER BY "displayOrder" ASC`,
    [cls.subjectId]
  );

  // Fetch comment options for all groups
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

  const groups = groupRows;

  // Fetch assignments with active pupils
  const { rows: assignmentRows } = await pool.query(
    `SELECT a.*,
            p."admissionNumber" as "pupil_admissionNumber",
            p."firstName" as "pupil_firstName",
            p."lastName" as "pupil_lastName",
            p.gender as "pupil_gender",
            p."isActive" as "pupil_isActive",
            p.form as "pupil_form"
     FROM "Assignment" a
     JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
     WHERE a."classId" = $1 AND p."isActive" = true`,
    [classId]
  );

  // Fetch PupilCodes for all assignments
  const assignmentIds = assignmentRows.map((r: any) => r.id);
  let pupilCodesByAssignment = new Map<string, any[]>();

  if (assignmentIds.length > 0) {
    const { rows: pupilCodeRows } = await pool.query(
      `SELECT * FROM "PupilCode" WHERE "assignmentId" = ANY($1::text[])`,
      [assignmentIds]
    );
    for (const pc of pupilCodeRows) {
      const arr = pupilCodesByAssignment.get(pc.assignmentId) ?? [];
      arr.push(pc);
      pupilCodesByAssignment.set(pc.assignmentId, arr);
    }
  }

  // Fetch CommonPupilCodes for all assignments
  let commonPupilCodesByAssignment = new Map<string, any[]>();

  if (assignmentIds.length > 0) {
    const { rows: commonPupilCodeRows } = await pool.query(
      `SELECT * FROM "CommonPupilCode" WHERE "assignmentId" = ANY($1::text[])`,
      [assignmentIds]
    );
    for (const cpc of commonPupilCodeRows) {
      const arr = commonPupilCodesByAssignment.get(cpc.assignmentId) ?? [];
      arr.push(cpc);
      commonPupilCodesByAssignment.set(cpc.assignmentId, arr);
    }
  }

  // Decrypt and assemble assignments
  const assignments = assignmentRows.map((row: any) => ({
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
    Pupil: {
      admissionNumber: row.pupil_admissionNumber,
      firstName: decrypt(row.pupil_firstName),
      lastName: decrypt(row.pupil_lastName),
      gender: row.pupil_gender,
      isActive: row.pupil_isActive,
      form: row.pupil_form,
    },
    PupilCode: pupilCodesByAssignment.get(row.id) ?? [],
    CommonPupilCode: commonPupilCodesByAssignment.get(row.id) ?? [],
  }));

  // Re-sort because database sort was on encrypted strings
  assignments.sort((a: any, b: any) =>
    a.Pupil.lastName.localeCompare(b.Pupil.lastName)
  );

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
  const subjectFormat = cls.Subject.commentFormat || null;

  // Split CCG groups into before-SCG and after-SCG based on template position
  let commonGroupsBefore = commonGroups as any[];
  let commonGroupsAfter: any[] = [];

  if (formatTemplate && formatTemplate.includes('<SCG>')) {
    const scgIdx = formatTemplate.indexOf('<SCG>');
    const namesAfterSCG = new Set([...formatTemplate.slice(scgIdx + 5).matchAll(/<([^>]+)>/g)].map((m: RegExpMatchArray) => m[1]));
    const posMap = new Map([...formatTemplate.matchAll(/<([^>]+)>/g)].map((m: RegExpMatchArray, i: number) => [m[1], i]));
    const byPos = (a: any, b: any) => (posMap.get(a.name) ?? 999) - (posMap.get(b.name) ?? 999);
    commonGroupsBefore = (commonGroups as any[]).filter((g: any) => !namesAfterSCG.has(g.name)).sort(byPos);
    commonGroupsAfter = (commonGroups as any[]).filter((g: any) => namesAfterSCG.has(g.name)).sort(byPos);
  }

  // Remove CCG columns that are overridden by a subject-specific CommentGroup with the same name
  const subjectCcgOverrideNames = new Set(groups.filter((g: any) => commonGroups.some((cg: any) => cg.name === g.name)).map((g: any) => g.name));
  commonGroupsBefore = commonGroupsBefore.filter((g: any) => !subjectCcgOverrideNames.has(g.name));
  commonGroupsAfter = commonGroupsAfter.filter((g: any) => !subjectCcgOverrideNames.has(g.name));

  return (
    <main className="flex-1 flex flex-col min-w-0 bg-background-light dark:bg-background-dark h-[calc(100vh-64px)] overflow-hidden">
        {/* Page Heading */}
        <div className="flex flex-wrap items-center justify-between gap-4 p-6 bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748]">
            <div className="flex flex-col gap-1">
                <div className="flex items-center gap-2">
                    <Link href="/" className="text-gray-500 hover:text-gray-900 transition-colors">
                        <ArrowLeft className="w-5 h-5" />
                    </Link>
                    <h1 className="text-[#111418] dark:text-white text-3xl font-black leading-tight tracking-[-0.033em]">
                        Class Matrix: {cls.name}
                    </h1>
                </div>
                <p className="text-[#617289] dark:text-gray-400 text-sm font-normal leading-normal">
                    {cls.Subject.code} • {assignments.length} Pupils
                </p>
            </div>
            <div className="flex gap-3">
                <button className="flex items-center justify-center rounded-lg h-10 px-4 bg-background-light dark:bg-[#2d3748] text-[#111418] dark:text-white text-sm font-bold tracking-[0.015em] border border-[#dbe0e6] dark:border-[#3a4454] hover:bg-gray-100 dark:hover:bg-[#3a4454] transition-colors shadow-sm">
                    <span className="material-symbols-outlined mr-2 text-lg">download</span>
                    Export CSV
                </button>
            </div>
        </div>

        {/* Table Content Container */}
        <div className="flex-1 overflow-auto p-6">
            <div className="bg-white dark:bg-[#1a222c] rounded-xl border border-[#dbe0e6] dark:border-[#2d3748] shadow-sm overflow-hidden h-full flex flex-col">
                <div className="overflow-auto relative h-full">
                    <table className="w-full text-left border-separate border-spacing-0">
                        <thead className="bg-white dark:bg-[#1a222c] sticky top-0 z-30">
                            <tr>
                                <th scope="col" className="sticky top-0 left-0 z-40 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] w-[240px] min-w-[240px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                    Pupil Name
                                </th>
                                <th scope="col" className="sticky top-0 left-[240px] z-40 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] w-[80px] min-w-[80px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                    Gender
                                </th>
                                <th scope="col" className="sticky top-0 left-[320px] z-40 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] w-[140px] min-w-[140px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
                                    Status
                                </th>
                                {/* CCG columns before SCG */}
                                {commonGroupsBefore.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-green-700 dark:text-green-400 text-xs font-bold uppercase tracking-wider bg-green-50/50 dark:bg-green-900/10 border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        <Tooltip content={g.name}>
                                            {g.name}
                                        </Tooltip>
                                    </th>
                                ))}
                                {/* Subject-specific group columns */}
                                {groups.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-[#111418] dark:text-white text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        <Tooltip content={g.name}>
                                            {g.name}
                                        </Tooltip>
                                    </th>
                                ))}
                                {/* CCG columns after SCG */}
                                {commonGroupsAfter.map((g: any) => (
                                    <th key={g.id} scope="col" className="sticky top-0 z-30 px-6 py-4 text-green-700 dark:text-green-400 text-xs font-bold uppercase tracking-wider bg-green-50/50 dark:bg-green-900/10 border-b border-[#e5e7eb] dark:border-[#2d3748] min-w-[200px]">
                                        <Tooltip content={g.name}>
                                            {g.name}
                                        </Tooltip>
                                    </th>
                                ))}
                                <th scope="col" className="sticky top-0 right-0 z-30 px-6 py-4 text-[#617289] dark:text-gray-400 text-xs font-bold uppercase tracking-wider bg-white dark:bg-[#1a222c] border-b border-[#e5e7eb] dark:border-[#2d3748] text-right min-w-[120px]">
                                    Actions
                                </th>
                            </tr>
                        </thead>
                        <tbody className="divide-y divide-[#e5e7eb] dark:divide-[#2d3748]">
                            {assignments.map((assignment: any) => (
                                <StudentMatrixRow
                                    key={assignment.id}
                                    assignment={assignment}
                                    groups={groups}
                                    subject={cls.Subject}
                                    classYear={cls.year}
                                    commonGroupsBefore={commonGroupsBefore}
                                    commonGroupsAfter={commonGroupsAfter}
                                    formatTemplate={formatTemplate}
                                    subjectFormat={subjectFormat}
                                />
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </main>
  );
}
