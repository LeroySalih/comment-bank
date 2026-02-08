'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { deleteCommonCommentGroup } from '@/lib/server-actions/ccg';
import CcgGroupForm from './ccg-group-form';
import CcgOptionForm from './ccg-option-form';
import CcgOptionItem from './ccg-option-item';
import ConfirmModal from '@/components/ConfirmModal';
import { parseComment } from '@/lib/utils';

type Option = {
  id: string;
  code: string;
  text: string;
  displayOrder: number;
};

type Group = {
  id: string;
  name: string;
  title: string;
  displayOrder: number;
  paragraphPosition: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommonCommentOption: Option[];
};

interface CcgGroupCardProps {
  group: Group;
  availableLinkedFields?: string[];
}

const SAMPLE_FIRST_NAME = 'Alex';
const SAMPLE_GENDER = 'M';
const SAMPLE_SUBJECT = 'Science';
const SAMPLE_YEAR = '7';
const SAMPLE_EOY = '7H';
const SAMPLE_TARGET = '7M';

function previewText(text: string): string {
  return parseComment(text, SAMPLE_FIRST_NAME, SAMPLE_GENDER, SAMPLE_SUBJECT, SAMPLE_YEAR, SAMPLE_EOY, SAMPLE_TARGET);
}

export default function CcgGroupCard({ group, availableLinkedFields = [] }: CcgGroupCardProps) {
  const router = useRouter();
  const [isEditing, setIsEditing] = useState(false);
  const [showAddOption, setShowAddOption] = useState(false);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [showAllPreviews, setShowAllPreviews] = useState(false);

  const handleDelete = async () => {
    setIsDeleting(true);
    try {
      const result = await deleteCommonCommentGroup(group.id);
      if (result.success) {
        router.refresh();
      } else {
        alert((result as any).error || 'Failed to delete group');
      }
    } catch {
      alert('Failed to delete group');
    }
    setIsDeleting(false);
    setShowDeleteModal(false);
  };

  if (isEditing) {
    return <CcgGroupForm group={group} availableLinkedFields={availableLinkedFields} onClose={() => setIsEditing(false)} />;
  }

  return (
    <div className="bg-white border border-gray-200 rounded-xl shadow-sm overflow-hidden">
      {/* Group Header */}
      <div className="px-6 py-4 bg-gray-50 border-b border-gray-200 flex items-center justify-between">
        <div>
          <h3 className="text-base font-bold text-gray-900">{group.name}</h3>
          {group.title !== group.name && (
            <p className="text-sm text-gray-500">{group.title}</p>
          )}
        </div>
        <div className="flex items-center gap-2">
          <span className="text-xs font-medium text-gray-500 bg-gray-200 px-2 py-1 rounded">
            {group.paragraphPosition.toUpperCase()}
          </span>
          {group.isLinked && (
            <span className="inline-flex items-center gap-1 text-xs font-medium text-purple-700 bg-purple-100 px-2 py-1 rounded">
              <span className="material-symbols-outlined text-sm">link</span>
              {group.linkedField}
            </span>
          )}
          <button
            onClick={() => setShowAllPreviews(p => !p)}
            className={`p-1.5 rounded transition-colors ${showAllPreviews ? 'text-primary bg-primary/10' : 'text-gray-500 hover:text-primary hover:bg-blue-50'}`}
            title={showAllPreviews ? 'Hide all previews' : 'Show all previews'}
          >
            <span className="material-symbols-outlined text-lg">preview</span>
          </button>
          <button
            onClick={() => setIsEditing(true)}
            className="p-1.5 text-gray-500 hover:text-primary hover:bg-blue-50 rounded transition-colors"
            title="Edit group"
          >
            <span className="material-symbols-outlined text-lg">edit</span>
          </button>
          <button
            onClick={() => setShowDeleteModal(true)}
            className="p-1.5 text-gray-500 hover:text-red-600 hover:bg-red-50 rounded transition-colors"
            title="Delete group"
          >
            <span className="material-symbols-outlined text-lg">delete</span>
          </button>
        </div>
      </div>

      {/* Preview All Banner */}
      {showAllPreviews && group.CommonCommentOption.length > 0 && (
        <div className="px-6 py-3 bg-blue-50/50 border-b border-blue-100">
          <p className="text-[10px] font-bold text-blue-400 uppercase tracking-wider mb-2">Preview (sample: Alex, Male, Science)</p>
          <div className="space-y-2">
            {group.CommonCommentOption.map(opt => (
              <div key={opt.id} className="flex gap-3 items-start">
                <span className="inline-flex items-center justify-center w-7 h-5 bg-primary/10 text-primary text-[10px] font-bold rounded flex-shrink-0 mt-0.5">
                  {opt.code}
                </span>
                <p className="text-sm text-gray-600 italic">{previewText(opt.text)}</p>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Options */}
      <div className="divide-y divide-gray-100">
        {group.CommonCommentOption.length === 0 ? (
          <div className="px-6 py-6 text-center text-sm text-gray-400">
            No options yet — add one below
          </div>
        ) : (
          group.CommonCommentOption.map(option => (
            <CcgOptionItem key={option.id} option={option} groupId={group.id} />
          ))
        )}
      </div>

      {/* Add Option */}
      <div className="px-6 py-3 bg-gray-50 border-t border-gray-200">
        {showAddOption ? (
          <CcgOptionForm groupId={group.id} onClose={() => setShowAddOption(false)} />
        ) : (
          <button
            onClick={() => setShowAddOption(true)}
            className="flex items-center gap-1 text-sm text-primary hover:text-blue-700 font-medium"
          >
            <span className="material-symbols-outlined text-lg">add</span>
            Add Option
          </button>
        )}
      </div>

      {/* Delete Confirmation */}
      <ConfirmModal
        isOpen={showDeleteModal}
        title="Delete Group?"
        message={`This will permanently delete "${group.name}" and all its options. This cannot be undone.`}
        confirmLabel={isDeleting ? 'Deleting...' : 'Delete Group'}
        cancelLabel="Cancel"
        confirmVariant="danger"
        onConfirm={handleDelete}
        onCancel={() => setShowDeleteModal(false)}
      />
    </div>
  );
}
