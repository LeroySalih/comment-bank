'use client';

import { useState } from 'react';
import Link from 'next/link';
import { useRouter } from 'next/navigation';
import QuickGroupSelector from '@/components/QuickGroupSelector';
import CopyCommentButton from '@/components/CopyCommentButton';
import CommentStatusBadge from '@/components/CommentStatusBadge';
import ConfirmModal from '@/components/ConfirmModal';
import { revertAssignmentComment, updateCommonAssignmentCode } from '@/app/actions';

type Option = {
  id: string;
  code: string;
  text: string;
};

type Group = {
  id: string;
  name: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommentOption: Option[];
};

type CommonCommentGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommonCommentOption: Option[];
};

type Pupil = {
  firstName: string;
  lastName: string;
  gender: string;
};

type PupilCode = {
  groupId: string;
  code: string | null;
};

type CommonPupilCode = {
  commonGroupId: string;
  code: string | null;
};

type Assignment = {
  id: string;
  Pupil: Pupil;
  PupilCode: PupilCode[];
  CommonPupilCode?: CommonPupilCode[];
  eoyLevel?: string | null;
  targetLevel?: string | null;
  checkStatus?: string;
  checkNote?: string | null;
  finalComment?: string | null;
  linkedData?: any;
};

type Subject = {
  title?: string | null;
  studiedComment?: string | null;
};

interface StudentMatrixRowProps {
  assignment: Assignment;
  groups: Group[];
  subject: Subject;
  classYear?: string | null;
  commonGroupsBefore?: CommonCommentGroup[];
  commonGroupsAfter?: CommonCommentGroup[];
  allCommonGroups?: CommonCommentGroup[];
  formatTemplate?: string;
  subjectFormat?: string | null;
}

function LinkedCodeBadge({ code, hasMatch }: { code: string | null; hasMatch: boolean }) {
  if (!code) {
    return (
      <span className="inline-flex items-center gap-1 text-xs text-gray-400 italic">
        <span className="material-symbols-outlined text-sm">link_off</span>
        No data
      </span>
    );
  }
  return (
    <span className={`inline-flex items-center gap-1 text-xs font-bold px-2.5 py-1.5 rounded-lg border ${
      hasMatch
        ? 'text-purple-700 bg-purple-50 border-purple-200'
        : 'text-amber-700 bg-amber-50 border-amber-200'
    }`}>
      <span className="material-symbols-outlined text-sm">{hasMatch ? 'link' : 'warning'}</span>
      {code}
    </span>
  );
}

export default function StudentMatrixRow({ assignment, groups, subject, classYear, commonGroupsBefore, commonGroupsAfter, allCommonGroups, formatTemplate, subjectFormat }: StudentMatrixRowProps) {
  const router = useRouter();

  // Track subject-specific selections locally
  const [selections, setSelections] = useState<Record<string, string | null>>(() => {
    const initial: Record<string, string | null> = {};
    assignment.PupilCode.forEach(pc => {
      initial[pc.groupId] = pc.code;
    });
    return initial;
  });

  // Track common group selections locally
  const [commonSelections, setCommonSelections] = useState<Record<string, string | null>>(() => {
    const initial: Record<string, string | null> = {};
    const linkedData = assignment.linkedData as Record<string, string> | null | undefined;
    [...(commonGroupsBefore || []), ...(commonGroupsAfter || [])].forEach(cg => {
      if (cg.isLinked && cg.linkedField && linkedData) {
        initial[cg.id] = linkedData[cg.linkedField] || null;
      } else {
        const cpc = (assignment.CommonPupilCode || []).find(c => c.commonGroupId === cg.id);
        initial[cg.id] = cpc?.code || null;
      }
    });
    return initial;
  });

  // Disable comment banks if comment has been manually edited (has finalComment)
  const [isReverted, setIsReverted] = useState(false);
  const commentBanksDisabled = !!assignment.finalComment && !isReverted;

  // Modal state for revert confirmation
  const [showRevertModal, setShowRevertModal] = useState(false);
  const [isReverting, setIsReverting] = useState(false);

  const handleSelectionChange = (groupId: string, code: string | null) => {
    if (commentBanksDisabled) return;
    setSelections(prev => ({
      ...prev,
      [groupId]: code
    }));
  };

  const handleCommonSelectionChange = (groupId: string, code: string | null) => {
    if (commentBanksDisabled) return;
    setCommonSelections(prev => ({
      ...prev,
      [groupId]: code
    }));
  };

  const handleRevert = async () => {
    setIsReverting(true);
    try {
      const result = await revertAssignmentComment(assignment.id);
      if (result.success) {
        setIsReverted(true);
        setShowRevertModal(false);
        router.refresh();
      } else {
        alert('Failed to revert: ' + result.error);
      }
    } catch (error) {
      alert('Failed to revert comment');
    }
    setIsReverting(false);
  };

  // Build the assignment object with current selections for CopyCommentButton
  const currentAssignment = {
    ...assignment,
    PupilCode: groups.map(g => ({
      groupId: g.id,
      code: selections[g.id] || null
    })),
    finalComment: assignment.finalComment,
    checkStatus: assignment.checkStatus,
    linkedData: assignment.linkedData
  };

  // Build current common pupil codes for CopyCommentButton
  const currentCommonPupilCodes = [...(commonGroupsBefore || []), ...(commonGroupsAfter || [])].map(g => ({
    commonGroupId: g.id,
    code: commonSelections[g.id] || null
  }));

  const contextForTooltip = {
    firstName: assignment.Pupil.firstName,
    gender: assignment.Pupil.gender,
    subjectTitle: subject.title || undefined,
    year: classYear || undefined,
    eoyLevel: assignment.eoyLevel,
    targetLevel: assignment.targetLevel
  };

  return (
    <tr className="group hover:bg-primary/5 dark:hover:bg-primary/10 transition-colors">
      <td className="sticky left-0 z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-blue-50 dark:group-hover:bg-[#1d2838] border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
        <span className="text-[#111418] dark:text-white text-sm font-semibold">
          {assignment.Pupil.lastName}, {assignment.Pupil.firstName}
        </span>
      </td>
      <td className="sticky left-[240px] z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-blue-50 dark:group-hover:bg-[#1d2838] border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
        <span className="text-sm text-[#617289] dark:text-gray-400">{assignment.Pupil.gender}</span>
      </td>
      <td className="sticky left-[320px] z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-blue-50 dark:group-hover:bg-[#1d2838] border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
        <div className="flex items-center gap-2">
          <CommentStatusBadge status={assignment.checkStatus || 'not_required'} showLabel={true} size="sm" />
          {commentBanksDisabled && (
            <button
              onClick={() => setShowRevertModal(true)}
              className="p-1 text-amber-600 hover:text-amber-700 hover:bg-amber-50 dark:hover:bg-amber-900/20 rounded transition-colors"
              title="Revert to generated comment"
            >
              <span className="material-symbols-outlined text-lg">undo</span>
            </button>
          )}
        </div>
      </td>
      <td className="sticky left-[460px] z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-blue-50 dark:group-hover:bg-[#1d2838] border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
        {assignment.eoyLevel ? (
          <span className="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-semibold bg-indigo-50 text-indigo-700 dark:bg-indigo-900/30 dark:text-indigo-300 border border-indigo-200 dark:border-indigo-800">
            {assignment.eoyLevel}
          </span>
        ) : (
          <span className="text-xs text-gray-400 italic">—</span>
        )}
      </td>
      <td className="sticky left-[560px] z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-blue-50 dark:group-hover:bg-[#1d2838] border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
        {assignment.targetLevel ? (
          <span className="inline-flex items-center px-2.5 py-1 rounded-md text-xs font-semibold bg-emerald-50 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800">
            {assignment.targetLevel}
          </span>
        ) : (
          <span className="text-xs text-gray-400 italic">—</span>
        )}
      </td>

      {/* CCG columns before SCG */}
      {(commonGroupsBefore || []).map((g) => {
        const currentCode = commonSelections[g.id] || null;
        if (g.isLinked) {
          const hasMatch = currentCode && g.CommonCommentOption.some(o => o.code === currentCode);
          return (
            <td key={g.id} className="px-6 py-4 whitespace-nowrap">
              <LinkedCodeBadge code={currentCode} hasMatch={!!hasMatch} />
            </td>
          );
        }
        return (
          <td key={g.id} className="px-6 py-4 whitespace-nowrap">
            <QuickGroupSelector
              assignmentId={assignment.id}
              groupId={g.id}
              currentCode={currentCode}
              options={g.CommonCommentOption}
              context={contextForTooltip}
              onSelectionChange={handleCommonSelectionChange}
              onCodeUpdate={updateCommonAssignmentCode}
              disabled={commentBanksDisabled}
            />
          </td>
        );
      })}

      {/* Subject-specific group columns */}
      {groups.map((g) => {
        const currentCode = selections[g.id] || null;
        return (
          <td key={g.id} className="px-6 py-4 whitespace-nowrap">
            <QuickGroupSelector
              assignmentId={assignment.id}
              groupId={g.id}
              currentCode={currentCode}
              options={g.CommentOption}
              context={contextForTooltip}
              onSelectionChange={handleSelectionChange}
              disabled={commentBanksDisabled}
            />
          </td>
        );
      })}

      {/* CCG columns after SCG */}
      {(commonGroupsAfter || []).map((g) => {
        const currentCode = commonSelections[g.id] || null;
        if (g.isLinked) {
          const hasMatch = currentCode && g.CommonCommentOption.some(o => o.code === currentCode);
          return (
            <td key={g.id} className="px-6 py-4 whitespace-nowrap">
              <LinkedCodeBadge code={currentCode} hasMatch={!!hasMatch} />
            </td>
          );
        }
        return (
          <td key={g.id} className="px-6 py-4 whitespace-nowrap">
            <QuickGroupSelector
              assignmentId={assignment.id}
              groupId={g.id}
              currentCode={currentCode}
              options={g.CommonCommentOption}
              context={contextForTooltip}
              onSelectionChange={handleCommonSelectionChange}
              onCodeUpdate={updateCommonAssignmentCode}
              disabled={commentBanksDisabled}
            />
          </td>
        );
      })}

      <td className="px-6 py-4 whitespace-nowrap text-right">
        <div className="flex items-center justify-end gap-3">
          <CopyCommentButton
            assignment={currentAssignment}
            subject={subject}
            groups={groups}
            commonGroups={allCommonGroups || [...(commonGroupsBefore || []), ...(commonGroupsAfter || [])]}
            commonPupilCodes={currentCommonPupilCodes}
            formatTemplate={formatTemplate}
            subjectFormat={subjectFormat}
          />
          <Link
            href={`/student/${assignment.id}`}
            className="text-primary hover:text-blue-700 text-sm font-bold transition-colors inline-flex items-center gap-1"
          >
            <span className="material-symbols-outlined text-lg">visibility</span>
            Preview
          </Link>
        </div>
      </td>

      {/* Revert Confirmation Modal - rendered via portal to document.body */}
      <ConfirmModal
        isOpen={showRevertModal}
        title="Revert to Generated Comment?"
        message={`This will discard the manual edits for ${assignment.Pupil.firstName} ${assignment.Pupil.lastName} and regenerate the comment from the selected codes. This action cannot be undone.`}
        confirmLabel={isReverting ? 'Reverting...' : 'Revert Comment'}
        cancelLabel="Keep Edits"
        confirmVariant="danger"
        onConfirm={handleRevert}
        onCancel={() => setShowRevertModal(false)}
      />
    </tr>
  );
}
