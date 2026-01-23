'use client';

import { useState } from 'react';
import { Copy, Check } from 'lucide-react';
import { generateComment } from '@/lib/comment-utils';

type MinimalOption = {
  id: string;
  code: string;
  text: string;
};

type MinimalGroup = {
  id: string;
  name: string;
  options: MinimalOption[];
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
  pupil: MinimalPupil;
  codes: MinimalPupilCode[];
  eoyLevel?: string | null;
  targetLevel?: string | null;
};

type MinimalSubject = {
  studiedComment?: string | null;
  subject?: string | null;
};

interface CopyCommentButtonProps {
  assignment: MinimalAssignment;
  subject: MinimalSubject;
  groups: MinimalGroup[];
}

export default function CopyCommentButton({ assignment, subject, groups }: CopyCommentButtonProps) {
  const [copied, setCopied] = useState(false);

  const handleCopy = (e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();

    const comment = generateComment(assignment, subject, groups);
    
    if (!comment) return;

    navigator.clipboard.writeText(comment);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <button
      onClick={handleCopy}
      className={`
        p-2 rounded-full transition-all duration-200
        ${copied 
          ? 'bg-green-100 text-green-600 hover:bg-green-200' 
          : 'text-gray-400 hover:text-blue-600 hover:bg-blue-50'
        }
      `}
      title="Copy Comment"
    >
      {copied ? (
        <Check className="w-4 h-4" />
      ) : (
        <Copy className="w-4 h-4" />
      )}
      <span className="sr-only">Copy Comment</span>
    </button>
  );
}
