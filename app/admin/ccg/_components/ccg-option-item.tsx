'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { updateCommonCommentOption, deleteCommonCommentOption } from '@/lib/server-actions/ccg';
import { parseComment } from '@/lib/utils';
import VariableAutocompleteInput from '@/components/VariableAutocompleteInput';

type Option = {
  id: string;
  code: string;
  text: string;
  displayOrder: number;
};

interface CcgOptionItemProps {
  option: Option;
  groupId: string;
}

// Sample values for preview
const SAMPLE_FIRST_NAME = 'Alex';
const SAMPLE_GENDER = 'M';
const SAMPLE_SUBJECT = 'Science';
const SAMPLE_YEAR = '7';
const SAMPLE_EOY = '7H';
const SAMPLE_TARGET = '7M';

function previewText(text: string): string {
  return parseComment(text, SAMPLE_FIRST_NAME, SAMPLE_GENDER, SAMPLE_SUBJECT, SAMPLE_YEAR, SAMPLE_EOY, SAMPLE_TARGET);
}

export default function CcgOptionItem({ option, groupId }: CcgOptionItemProps) {
  const router = useRouter();
  const [isEditing, setIsEditing] = useState(false);
  const [loading, setLoading] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [editText, setEditText] = useState(option.text);

  const handleUpdate = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setLoading(true);

    const formData = new FormData(e.currentTarget);
    const code = formData.get('code') as string;

    try {
      const result = await updateCommonCommentOption(option.id, code, editText);
      if (result.success) {
        setIsEditing(false);
        router.refresh();
      } else {
        alert((result as any).error || 'Failed to update option');
      }
    } catch (err) {
      console.error('Failed to update option:', err);
      alert('Failed to update option: ' + (err instanceof Error ? err.message : String(err)));
    }

    setLoading(false);
  };

  const handleDelete = async () => {
    if (!confirm(`Delete option "${option.code}"?`)) return;
    setLoading(true);

    try {
      const result = await deleteCommonCommentOption(option.id);
      if (result.success) {
        router.refresh();
      } else {
        alert((result as any).error || 'Failed to delete option');
      }
    } catch (err) {
      console.error('Failed to delete option:', err);
      alert('Failed to delete option: ' + (err instanceof Error ? err.message : String(err)));
    }

    setLoading(false);
  };

  if (isEditing) {
    return (
      <div className="px-6 py-3 space-y-2">
        <form onSubmit={handleUpdate} className="flex gap-3 items-center">
          <div className="w-20">
            <input
              name="code"
              type="text"
              defaultValue={option.code}
              required
              className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
            />
          </div>
          <VariableAutocompleteInput
            name="text"
            value={editText}
            onChange={setEditText}
            required
            autoFocus
          />
          <div className="flex gap-2 flex-shrink-0">
            <button
              type="submit"
              disabled={loading}
              className="px-3 py-1.5 text-xs font-medium text-white bg-primary rounded hover:bg-blue-700 disabled:opacity-50"
            >
              {loading ? '...' : 'Save'}
            </button>
            <button
              type="button"
              onClick={() => { setIsEditing(false); setEditText(option.text); }}
              className="px-3 py-1.5 text-xs font-medium text-gray-600 bg-gray-100 rounded hover:bg-gray-200"
            >
              Cancel
            </button>
          </div>
        </form>
        {/* Live preview while editing */}
        {editText && (
          <div className="ml-20 pl-3 border-l-2 border-primary/20">
            <p className="text-xs text-gray-400 mb-0.5">Preview:</p>
            <p className="text-sm text-gray-600 italic">{previewText(editText)}</p>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="px-6 py-3 group hover:bg-gray-50">
      <div className="flex items-center gap-4">
        <span className="inline-flex items-center justify-center w-8 h-6 bg-primary/10 text-primary text-xs font-bold rounded flex-shrink-0">
          {option.code}
        </span>
        <div className="flex-1 min-w-0">
          <p className="text-sm text-gray-700 truncate">{option.text}</p>
          {showPreview && (
            <p className="text-xs text-gray-500 italic mt-1 border-l-2 border-primary/20 pl-2">
              {previewText(option.text)}
            </p>
          )}
        </div>
        <div className="flex gap-1 opacity-0 group-hover:opacity-100 transition-opacity flex-shrink-0">
          <button
            onClick={() => setShowPreview(p => !p)}
            className={`p-1 rounded transition-colors ${showPreview ? 'text-primary bg-primary/10' : 'text-gray-400 hover:text-primary'}`}
            title={showPreview ? 'Hide preview' : 'Show preview'}
          >
            <span className="material-symbols-outlined text-base">preview</span>
          </button>
          <button
            onClick={() => { setIsEditing(true); setEditText(option.text); }}
            className="p-1 text-gray-400 hover:text-primary rounded transition-colors"
            title="Edit"
          >
            <span className="material-symbols-outlined text-base">edit</span>
          </button>
          <button
            onClick={handleDelete}
            disabled={loading}
            className="p-1 text-gray-400 hover:text-red-600 rounded transition-colors disabled:opacity-50"
            title="Delete"
          >
            <span className="material-symbols-outlined text-base">delete</span>
          </button>
        </div>
      </div>
    </div>
  );
}
