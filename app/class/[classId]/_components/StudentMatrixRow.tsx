'use client';

import { useState } from 'react';
import Link from 'next/link';
import { useRouter } from 'next/navigation';
import QuickGroupSelector from '@/components/QuickGroupSelector';
import CopyCommentButton from '@/components/CopyCommentButton';
import CommentStatusBadge from '@/components/CommentStatusBadge';
import ConfirmModal from '@/components/ConfirmModal';
import { revertAssignmentComment } from '@/app/actions';

type Option = {
  id: string;
  code: string;
  text: string;
};

type Group = {
  id: string;
  name: string;
  CommentOption: Option[];
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

type Assignment = {
  id: string;
  Pupil: Pupil;
  PupilCode: PupilCode[];
  eoyLevel?: string | null;
  targetLevel?: string | null;
  checkStatus?: string;
  checkNote?: string | null;
  finalComment?: string | null;
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
}

export default function StudentMatrixRow({ assignment, groups, subject, classYear }: StudentMatrixRowProps) {
  const router = useRouter();

  // Track selections locally so CopyCommentButton updates immediately
  const [selections, setSelections] = useState<Record<string, string | null>>(() => {
    const initial: Record<string, string | null> = {};
    assignment.PupilCode.forEach(pc => {
      initial[pc.groupId] = pc.code;
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
    if (commentBanksDisabled) return; // Don't allow changes if disabled
    setSelections(prev => ({
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
    // Pass through the saved finalComment and checkStatus
    finalComment: assignment.finalComment,
    checkStatus: assignment.checkStatus
  };

  return (
    <tr className="group hover:bg-primary/5 dark:hover:bg-primary/10 transition-colors">
      <td className="sticky left-0 z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-primary/5 dark:group-hover:bg-gray-800/50 border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
        <span className="text-[#111418] dark:text-white text-sm font-semibold">
          {assignment.Pupil.lastName}, {assignment.Pupil.firstName}
        </span>
      </td>
      <td className="sticky left-[240px] z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-primary/5 dark:group-hover:bg-gray-800/50 border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
        <span className="text-sm text-[#617289] dark:text-gray-400">{assignment.Pupil.gender}</span>
      </td>
      <td className="sticky left-[320px] z-20 px-6 py-4 whitespace-nowrap bg-white dark:bg-[#1a222c] group-hover:bg-primary/5 dark:group-hover:bg-gray-800/50 border-b border-gray-100 dark:border-gray-800 shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]">
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
      {groups.map((g) => {
        const currentCode = selections[g.id] || null;
        return (
          <td key={g.id} className="px-6 py-4 whitespace-nowrap">
            <QuickGroupSelector
              assignmentId={assignment.id}
              groupId={g.id}
              currentCode={currentCode}
              options={g.CommentOption}
              context={{
                firstName: assignment.Pupil.firstName,
                gender: assignment.Pupil.gender,
                subjectTitle: subject.title || undefined,
                year: classYear || undefined,
                eoyLevel: assignment.eoyLevel,
                targetLevel: assignment.targetLevel
              }}
              onSelectionChange={handleSelectionChange}
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
