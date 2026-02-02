'use client';

import { useState } from 'react';
import { updateAssignmentCode } from '@/app/actions';
import { generateComment } from '@/lib/comment-utils';

type Option = {
  id: string;
  code: string;
  text: string;
};

type CommentGroup = {
  id: string;
  name: string;
  CommentOption: Option[];
};

type Subject = {
  code: string;
  title?: string | null;
  studiedComment?: string | null;
  CommentGroup: CommentGroup[];
};

type Class = {
  id: string;
  year?: string | null;
  Subject: Subject;
};

type PupilCode = {
  id: string;
  groupId: string;
  code: string | null;
  CommentGroup: CommentGroup;
};

type Assignment = {
  id: string;
  eoyLevel?: string | null;
  targetLevel?: string | null;
  finalComment?: string | null;
  Class: Class;
  PupilCode: PupilCode[];
};

type Pupil = {
  admissionNumber: string;
  firstName: string;
  lastName: string;
  gender: string;
  Assignment: Assignment[];
};

interface FormPupilRowProps {
  pupil: Pupil;
  userAssignedClassIds: string[];
}

export default function FormPupilRow({ pupil, userAssignedClassIds }: FormPupilRowProps) {
  // Track selections for each assignment
  const [selections, setSelections] = useState<Record<string, Record<string, string | null>>>(() => {
    const initial: Record<string, Record<string, string | null>> = {};
    pupil.Assignment.forEach(assignment => {
      initial[assignment.id] = {};
      assignment.PupilCode.forEach(pc => {
        initial[assignment.id][pc.groupId] = pc.code;
      });
    });
    return initial;
  });

  const [loadingStates, setLoadingStates] = useState<Record<string, boolean>>({});

  const handleCodeChange = async (assignmentId: string, groupId: string, newCode: string | null) => {
    const key = `${assignmentId}-${groupId}`;
    setLoadingStates(prev => ({ ...prev, [key]: true }));

    // Update local state immediately
    setSelections(prev => ({
      ...prev,
      [assignmentId]: {
        ...prev[assignmentId],
        [groupId]: newCode
      }
    }));

    // Save to database
    await updateAssignmentCode(assignmentId, groupId, newCode);
    setLoadingStates(prev => ({ ...prev, [key]: false }));
  };

  return (
    <div className="grid grid-cols-[200px_1fr_1fr] gap-4 px-4 py-3 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">
      {/* Pupil Name */}
      <div className="flex flex-col justify-center">
        <span className="font-semibold text-slate-900 dark:text-white">
          {pupil.lastName}, {pupil.firstName}
        </span>
        <span className="text-xs text-slate-500 dark:text-slate-400">
          {pupil.gender}
        </span>
      </div>

      {/* Subjects & Codes */}
      <div className="flex flex-col gap-2">
        {pupil.Assignment.length === 0 ? (
          <span className="text-slate-400 text-sm italic">No assignments</span>
        ) : (
          pupil.Assignment.map((assignment) => {
            const canEdit = userAssignedClassIds.includes(assignment.Class.id);
            const groups = assignment.Class.Subject.CommentGroup;

            return (
              <div key={assignment.id} className="flex items-start gap-2 text-sm">
                <span className="font-medium text-slate-700 dark:text-slate-300 min-w-[80px] pt-1">
                  {assignment.Class.Subject.code}
                </span>
                <div className="flex flex-wrap gap-1">
                  {groups.map((group) => {
                    const currentCode = selections[assignment.id]?.[group.id] || null;
                    const loadingKey = `${assignment.id}-${group.id}`;
                    const isLoading = loadingStates[loadingKey];

                    return (
                      <div key={group.id} className="flex items-center gap-1">
                        <span className="text-xs text-slate-500 dark:text-slate-400">{group.name}:</span>
                        <select
                          value={currentCode || ''}
                          onChange={(e) => handleCodeChange(assignment.id, group.id, e.target.value || null)}
                          disabled={!canEdit || isLoading}
                          className={`text-xs px-1.5 py-0.5 rounded border ${
                            canEdit
                              ? 'border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 cursor-pointer hover:border-blue-400'
                              : 'border-slate-200 dark:border-slate-700 bg-slate-100 dark:bg-slate-800/50 cursor-not-allowed'
                          } text-slate-700 dark:text-slate-300 focus:outline-none focus:ring-1 focus:ring-blue-500 ${
                            isLoading ? 'opacity-50' : ''
                          }`}
                        >
                          <option value="">—</option>
                          {group.CommentOption.map((opt) => (
                            <option key={opt.id} value={opt.code}>
                              {opt.code}
                            </option>
                          ))}
                        </select>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })
        )}
      </div>

      {/* Comments */}
      <div className="flex flex-col gap-1.5 text-sm">
        {pupil.Assignment.length === 0 ? (
          <span className="text-slate-400 italic">—</span>
        ) : (
          pupil.Assignment.map((assignment) => {
            // Build current PupilCode array from selections
            const currentPupilCodes = assignment.Class.Subject.CommentGroup.map(g => ({
              groupId: g.id,
              code: selections[assignment.id]?.[g.id] || null
            }));

            // Generate comment with variables replaced
            const generatedComment = generateComment(
              {
                id: assignment.id,
                Pupil: {
                  firstName: pupil.firstName,
                  lastName: pupil.lastName,
                  gender: pupil.gender
                },
                PupilCode: currentPupilCodes,
                eoyLevel: assignment.eoyLevel,
                targetLevel: assignment.targetLevel
              },
              assignment.Class.Subject,
              assignment.Class.Subject.CommentGroup,
              assignment.Class
            );

            return (
              <div key={assignment.id} className="flex items-start gap-2">
                <span className="font-medium text-slate-500 dark:text-slate-400 min-w-[80px] text-xs">
                  {assignment.Class.Subject.code}:
                </span>
                {generatedComment ? (
                  <span className="text-slate-600 dark:text-slate-300 text-xs leading-relaxed">
                    {generatedComment}
                  </span>
                ) : (
                  <span className="text-slate-400 text-xs">—</span>
                )}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}
