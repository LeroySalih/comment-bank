'use client';

import { useState } from 'react';
import { updateAssignmentCode } from '@/app/actions';
import Tooltip from './Tooltip';
import { parseComment } from '@/lib/utils';

type Option = {
  id: string;
  code: string;
  text: string;
};

interface QuickGroupSelectorProps {
  assignmentId: string;
  groupId: string;
  currentCode: string | null | undefined;
  options: Option[];
  context?: {
    firstName: string;
    gender: string;
    subjectTitle?: string;
    year?: string | null;
    eoyLevel?: string | null;
    targetLevel?: string | null;
  };
  onSelectionChange?: (groupId: string, code: string | null) => void;
  onCodeUpdate?: (assignmentId: string, groupId: string, code: string | null) => Promise<any>;
  disabled?: boolean;
}

export default function QuickGroupSelector({
  assignmentId,
  groupId,
  currentCode,
  options,
  context,
  onSelectionChange,
  onCodeUpdate,
  disabled = false
}: QuickGroupSelectorProps) {
  const [val, setVal] = useState(currentCode || "");
  const [loading, setLoading] = useState(false);

  const updateFn = onCodeUpdate || updateAssignmentCode;

  const handleChange = async (newValue: string | null) => {
    if (disabled) return;

    setVal(newValue || "");
    setLoading(true);

    onSelectionChange?.(groupId, newValue);

    await updateFn(assignmentId, groupId, newValue);
    setLoading(false);
  };

  return (
    <div className={`flex rounded-lg overflow-hidden border border-[#dbe0e6] dark:border-[#3a4454] w-fit bg-gray-50 dark:bg-[#101822] divide-x divide-[#dbe0e6] dark:divide-[#3a4454] ${disabled ? 'opacity-50' : ''}`}>
      {options.map((opt) => {
        const isActive = val === opt.code;

        let tooltipContent = opt.text;
        if (disabled) {
            tooltipContent = 'Locked - comment has been manually edited';
        } else if (context) {
            tooltipContent = parseComment(
                opt.text,
                context.firstName,
                context.gender,
                context.subjectTitle,
                context.year,
                context.eoyLevel,
                context.targetLevel
            );
        }

        return (
          <Tooltip
            key={opt.id}
            content={<div className="font-normal"><span className="font-bold mb-1 block">{opt.code}</span>{tooltipContent}</div>}
            maxWidth="max-w-xs"
          >
            <button
                onClick={() => val === opt.code ? handleChange("") : handleChange(opt.code)}
                disabled={loading || disabled}
                className={`px-3 py-1.5 text-xs font-bold transition-colors min-w-[32px]
                ${isActive
                    ? 'bg-primary text-white'
                    : 'text-[#617289] dark:text-gray-400 hover:bg-gray-100 dark:hover:bg-[#2d3748]'}
                ${loading || disabled ? 'opacity-50 cursor-not-allowed' : ''}
                `}
            >
                {opt.code}
            </button>
          </Tooltip>
        );
      })}
    </div>
  );
}
