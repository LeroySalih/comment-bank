'use client';

import { useState } from 'react';
import { Copy, Check } from 'lucide-react';
import { generateComment } from '@/lib/comment-utils';
import Tooltip from './Tooltip';

type MinimalOption = {
  id: string;
  code: string;
  text: string;
};

type MinimalGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommentOption: MinimalOption[];
};

type MinimalPupil = {
  firstName: string;
  lastName: string;
  gender: string;
};

type MinimalPupilCode = {
  groupId: string;
  code: string | null;
};

type MinimalCommonPupilCode = {
  commonGroupId: string;
  code: string | null;
};

type MinimalCommonGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommonCommentOption: MinimalOption[];
};

type MinimalAssignment = {
  id: string;
  Pupil: MinimalPupil;
  PupilCode: MinimalPupilCode[];
  eoyLevel?: string | null;
  targetLevel?: string | null;
  finalComment?: string | null;
  checkStatus?: string;
  linkedData?: any;
};

type MinimalSubject = {
  studiedComment?: string | null;
  title?: string | null;
};

interface CopyCommentButtonProps {
  assignment: MinimalAssignment;
  subject: MinimalSubject;
  groups: MinimalGroup[];
  commonGroups?: MinimalCommonGroup[];
  commonPupilCodes?: MinimalCommonPupilCode[];
  formatTemplate?: string;
  subjectFormat?: string | null;
}

export default function CopyCommentButton({ assignment, subject, groups, commonGroups, commonPupilCodes, formatTemplate, subjectFormat }: CopyCommentButtonProps) {
  const [copied, setCopied] = useState(false);

  // Check if all subject groups have a selected code
  const subjectComplete = groups.every(g => {
    if (g.isLinked && g.linkedField) {
      const code = (assignment.linkedData as Record<string, string> | null)?.[g.linkedField];
      return code && g.CommentOption.some(o => o.code === code);
    }
    return assignment.PupilCode?.some(c => c.groupId === g.id && c.code);
  });

  // Check if all CCGs have a selected code
  const ccgComplete = !commonGroups || commonGroups.length === 0 || commonGroups.every(g => {
    if (g.isLinked && g.linkedField) {
      const code = (assignment.linkedData as Record<string, string> | null)?.[g.linkedField];
      return code && g.CommonCommentOption.some(o => o.code === code);
    }
    return commonPupilCodes?.some(c => c.commonGroupId === g.id && c.code);
  });

  const isComplete = subjectComplete && ccgComplete;

  // Check if comment needs review or is rejected
  const isPendingReview = assignment.checkStatus === 'required_check' || assignment.checkStatus === 'checked_rejected';

  const isDisabled = !isComplete || isPendingReview;

  const handleCopy = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();

    if (isDisabled) return;

    const comment = assignment.finalComment || generateComment(
      assignment, subject, groups, undefined,
      commonGroups, commonPupilCodes, formatTemplate, subjectFormat
    );

    if (!comment) return;

    navigator.clipboard.writeText(comment);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const [preview, setPreview] = useState<string | null>(null);

  const handleMouseEnter = () => {
    if (!isComplete) {
      setPreview("Please select options for all groups first");
      return;
    }
    if (isPendingReview) {
      setPreview("Comment must be approved before copying");
      return;
    }
    const comment = assignment.finalComment || generateComment(
      assignment, subject, groups, undefined,
      commonGroups, commonPupilCodes, formatTemplate, subjectFormat
    );
    setPreview(comment || "No comment generated");
  };

  return (
    <Tooltip
        content={
          <div className={`font-normal whitespace-pre-wrap ${isDisabled ? 'text-red-500' : ''}`}>
            {preview}
          </div>
        }
        maxWidth="max-w-md"
    >
        <button
        onMouseEnter={handleMouseEnter}
        onClick={handleCopy}
        disabled={isDisabled}
        className={`
            p-2 rounded-full transition-all duration-200
            ${!isComplete
              ? 'bg-red-50 text-red-500 cursor-not-allowed'
              : isPendingReview
                ? 'bg-amber-50 text-amber-500 cursor-not-allowed'
                : copied
                  ? 'bg-green-100 text-green-600 hover:bg-green-200'
                  : 'text-green-600 hover:bg-green-50'
            }
        `}
        title={!isComplete ? "Incomplete selection" : isPendingReview ? "Comment must be approved" : "Copy Comment"}
        >
        {copied ? (
            <Check className="w-4 h-4" />
        ) : (
            <Copy className="w-4 h-4" />
        )}
        <span className="sr-only">Copy Comment</span>
        </button>
    </Tooltip>
  );
}
