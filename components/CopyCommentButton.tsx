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

type MinimalStudent = {
  id: string;
  firstName: string;
  lastName: string;
  gender: string;
  wpCode?: string | null;
  thCode?: string | null;
  psCode?: string | null;
  oaCode?: string | null;
};

type MinimalCourse = {
  studiedComment?: string | null;
};

interface CopyCommentButtonProps {
  student: MinimalStudent;
  course: MinimalCourse;
  groups: MinimalGroup[];
}

export default function CopyCommentButton({ student, course, groups }: CopyCommentButtonProps) {
  const [copied, setCopied] = useState(false);

  const handleCopy = (e: React.MouseEvent) => {
    e.preventDefault(); // Prevent navigation if placed inside a Link (though it shouldn't be)
    e.stopPropagation();

    const comment = generateComment(student, course, groups);
    
    if (!comment) return; // Nothing to copy

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
