'use client';

import { useState } from 'react';
import { useClassMatrix } from './ClassMatrixContext';

type Option = {
  id: string;
  code: string;
  text: string;
};

type Group = {
  id: string;
  name: string;
  isLinked?: boolean;
  CommentOption: Option[];
};

type CommonCommentGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  CommonCommentOption: Option[];
};

interface BulkCodeButtonRowProps {
  groups: Group[];
  commonGroupsBefore: CommonCommentGroup[];
  commonGroupsAfter: CommonCommentGroup[];
}

function BulkButton({
  groupId,
  code,
  groupType,
}: {
  groupId: string;
  code: string;
  groupType: 'subject' | 'common';
}) {
  const { applyBulkCode } = useClassMatrix();
  const [loading, setLoading] = useState(false);

  const handleClick = async () => {
    setLoading(true);
    await applyBulkCode(groupId, code, groupType);
    setLoading(false);
  };

  return (
    <button
      onClick={handleClick}
      disabled={loading}
      title={`Set all unset pupils to ${code}`}
      className={`px-2.5 py-1 text-xs font-bold rounded border transition-colors
        border-[#dbe0e6] dark:border-[#3a4454]
        text-[#617289] dark:text-gray-400
        hover:bg-gray-100 dark:hover:bg-[#2d3748]
        disabled:opacity-50 disabled:cursor-not-allowed`}
    >
      {loading ? '…' : code}
    </button>
  );
}

export default function BulkCodeButtonRow({
  groups,
  commonGroupsBefore,
  commonGroupsAfter,
}: BulkCodeButtonRowProps) {
  const emptyCellClass =
    'sticky top-[57px] z-40 px-6 py-2 bg-gray-50 dark:bg-[#151d28] border-b border-[#e5e7eb] dark:border-[#2d3748]';

  return (
    <tr>
      {/* Fixed columns — empty */}
      <th className={`${emptyCellClass} left-0 w-[240px] min-w-[240px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[240px] w-[80px] min-w-[80px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[320px] w-[140px] min-w-[140px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[460px] w-[100px] min-w-[100px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[560px] w-[100px] min-w-[100px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />

      {/* CCG before SCG */}
      {commonGroupsBefore.map((g) => (
        <th key={g.id} className={`${emptyCellClass} min-w-[200px]`}>
          {!g.isLinked && (
            <div className="flex gap-1">
              {g.CommonCommentOption.map((opt) => (
                <BulkButton key={opt.id} groupId={g.id} code={opt.code} groupType="common" />
              ))}
            </div>
          )}
        </th>
      ))}

      {/* Subject-specific groups */}
      {groups.map((g) => (
        <th key={g.id} className={`${emptyCellClass} min-w-[200px]`}>
          {!g.isLinked && (
            <div className="flex gap-1">
              {g.CommentOption.map((opt) => (
                <BulkButton key={opt.id} groupId={g.id} code={opt.code} groupType="subject" />
              ))}
            </div>
          )}
        </th>
      ))}

      {/* CCG after SCG */}
      {commonGroupsAfter.map((g) => (
        <th key={g.id} className={`${emptyCellClass} min-w-[200px]`}>
          {!g.isLinked && (
            <div className="flex gap-1">
              {g.CommonCommentOption.map((opt) => (
                <BulkButton key={opt.id} groupId={g.id} code={opt.code} groupType="common" />
              ))}
            </div>
          )}
        </th>
      ))}

      {/* Actions column — empty */}
      <th className={`${emptyCellClass} right-0 min-w-[120px]`} />
    </tr>
  );
}
