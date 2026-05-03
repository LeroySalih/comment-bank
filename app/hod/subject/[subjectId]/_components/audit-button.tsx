'use client';

import { useState } from 'react';
import AuditModal from '@/components/AuditModal';

interface AuditButtonProps {
  subjectId: string;
  subjectTitle: string;
}

export function AuditButton({ subjectId, subjectTitle }: AuditButtonProps) {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <>
      <button
        onClick={() => setIsOpen(true)}
        className="flex items-center gap-2 px-4 py-2 bg-indigo-50 dark:bg-indigo-900/30 text-indigo-700 dark:text-indigo-400 rounded-lg hover:bg-indigo-100 dark:hover:bg-indigo-900/50 transition-colors font-medium text-sm"
      >
        <span className="material-symbols-outlined text-lg">fact_check</span>
        Audit Comments
      </button>
      <AuditModal
        subjectId={subjectId}
        subjectTitle={subjectTitle}
        isOpen={isOpen}
        onClose={() => setIsOpen(false)}
      />
    </>
  );
}
