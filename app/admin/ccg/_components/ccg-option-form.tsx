'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { createCommonCommentOption } from '@/lib/server-actions/ccg';
import { parseComment } from '@/lib/utils';
import VariableAutocompleteInput from '@/components/VariableAutocompleteInput';

interface CcgOptionFormProps {
  groupId: string;
  onClose: () => void;
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

export default function CcgOptionForm({ groupId, onClose }: CcgOptionFormProps) {
  const router = useRouter();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [text, setText] = useState('');

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setLoading(true);
    setError('');

    const formData = new FormData(e.currentTarget);
    const code = formData.get('code') as string;

    try {
      const result = await createCommonCommentOption(groupId, code, text);
      if (result.success) {
        onClose();
        router.refresh();
      } else {
        setError((result as any).error || 'Failed to create option');
      }
    } catch (err) {
      console.error('Failed to create option:', err);
      setError('An unexpected error occurred: ' + (err instanceof Error ? err.message : String(err)));
    }

    setLoading(false);
  };

  return (
    <div className="space-y-3">
      {error && (
        <div className="p-2 bg-red-50 border border-red-200 rounded text-xs text-red-700">
          {error}
        </div>
      )}

      <form onSubmit={handleSubmit} className="flex gap-3">
        <div className="w-20">
          <input
            name="code"
            type="text"
            required
            className="w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
            placeholder="Code"
          />
        </div>
        <VariableAutocompleteInput
          name="text"
          value={text}
          onChange={setText}
          required
          placeholder="Comment text — type < for variables"
        />
        <div className="flex gap-2 flex-shrink-0">
          <button
            type="submit"
            disabled={loading}
            className="px-3 py-1.5 text-xs font-medium text-white bg-primary rounded hover:bg-blue-700 disabled:opacity-50"
          >
            {loading ? '...' : 'Add'}
          </button>
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-medium text-gray-600 bg-gray-100 rounded hover:bg-gray-200"
          >
            Cancel
          </button>
        </div>
      </form>

      {/* Live preview */}
      {text && (
        <div className="ml-20 pl-3 border-l-2 border-primary/20">
          <p className="text-xs text-gray-400 mb-0.5">Preview:</p>
          <p className="text-sm text-gray-600 italic">{previewText(text)}</p>
        </div>
      )}
    </div>
  );
}
