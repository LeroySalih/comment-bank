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

type MinimalAssignment = {
  id: string;
  Pupil: MinimalPupil;
  PupilCode: MinimalPupilCode[];
  eoyLevel?: string | null;
  targetLevel?: string | null;
  finalComment?: string | null;
  checkStatus?: string;
};

type MinimalSubject = {
  studiedComment?: string | null;
  title?: string | null;
};

interface CopyCommentButtonProps {
  assignment: MinimalAssignment;
  subject: MinimalSubject;
  groups: MinimalGroup[];
}

export default function CopyCommentButton({ assignment, subject, groups }: CopyCommentButtonProps) {
  const [copied, setCopied] = useState(false);

  // Check if all groups have a selected code
  const isComplete = groups.every(g =>
    assignment.PupilCode?.some(c => c.groupId === g.id && c.code)
  );

  // Check if comment needs review or is rejected
  const isPendingReview = assignment.checkStatus === 'required_check' || assignment.checkStatus === 'checked_rejected';

  // Button is disabled if incomplete OR pending review
  const isDisabled = !isComplete || isPendingReview;

  const handleCopy = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();

    if (isDisabled) return;

    // Use saved finalComment if available, otherwise generate from codes
    const comment = assignment.finalComment || generateComment(assignment, subject, groups);
    
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
    // Use saved finalComment if available, otherwise generate from codes
    const comment = assignment.finalComment || generateComment(assignment, subject, groups);
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
