'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { createCommonCommentGroup, updateCommonCommentGroup } from '@/lib/server-actions/ccg';

interface CcgGroupFormProps {
  group?: {
    id: string;
    name: string;
    title: string;
    paragraphPosition: string;
    isLinked?: boolean;
    linkedField?: string | null;
  };
  availableLinkedFields?: string[];
  onClose: () => void;
}

export default function CcgGroupForm({ group, availableLinkedFields = [], onClose }: CcgGroupFormProps) {
  const router = useRouter();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [isLinked, setIsLinked] = useState(group?.isLinked ?? false);
  const [linkedField, setLinkedField] = useState(group?.linkedField ?? '');
  const isEditing = !!group;

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    setLoading(true);
    setError('');

    const formData = new FormData(e.currentTarget);
    const name = formData.get('name') as string;
    const title = formData.get('title') as string;
    const paragraphPosition = formData.get('paragraphPosition') as string;

    try {
      const result = isEditing
        ? await updateCommonCommentGroup(group.id, name, title, paragraphPosition, isLinked, isLinked ? linkedField : null)
        : await createCommonCommentGroup(name, title, paragraphPosition, isLinked, isLinked ? linkedField : null);

      if (result.success) {
        onClose();
        router.refresh();
      } else {
        setError((result as any).error || 'Failed to save group');
      }
    } catch (err) {
      setError('An unexpected error occurred');
    }

    setLoading(false);
  };

  return (
    <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm">
      <h3 className="text-lg font-bold mb-4">{isEditing ? 'Edit Group' : 'Create New Group'}</h3>

      {error && (
        <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">
          {error}
        </div>
      )}

      <form onSubmit={handleSubmit} className="space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Name</label>
            <input
              name="name"
              type="text"
              defaultValue={group?.name || ''}
              required
              className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
              placeholder="e.g. Effort"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Title</label>
            <input
              name="title"
              type="text"
              defaultValue={group?.title || ''}
              required
              className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
              placeholder="e.g. Effort & Engagement"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Paragraph Position</label>
            <select
              name="paragraphPosition"
              defaultValue={group?.paragraphPosition || 'p1'}
              required
              className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
            >
              <option value="p1">P1 — Introduction</option>
              <option value="p2">P2 — General Attributes</option>
              <option value="p4">P4 — Conclusion</option>
            </select>
          </div>
        </div>

        {/* Linked Group Toggle */}
        <div className="border-t border-gray-200 pt-4">
          <div className="flex items-center gap-3">
            <label className="relative inline-flex items-center cursor-pointer">
              <input
                type="checkbox"
                checked={isLinked}
                onChange={(e) => setIsLinked(e.target.checked)}
                className="sr-only peer"
              />
              <div className="w-9 h-5 bg-gray-200 peer-focus:outline-none peer-focus:ring-2 peer-focus:ring-primary/20 rounded-full peer peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-4 after:w-4 after:transition-all peer-checked:bg-primary"></div>
            </label>
            <div>
              <span className="text-sm font-medium text-gray-700">Linked Group</span>
              <p className="text-xs text-gray-500">Auto-populated from uploaded student data</p>
            </div>
          </div>

          {isLinked && (
            <div className="mt-3 ml-12">
              <label className="block text-sm font-medium text-gray-700 mb-1">Linked Field</label>
              {availableLinkedFields.length > 0 ? (
                <select
                  value={linkedField}
                  onChange={(e) => setLinkedField(e.target.value)}
                  required={isLinked}
                  className="w-full max-w-xs px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
                >
                  <option value="">Select a field...</option>
                  {availableLinkedFields.map(f => (
                    <option key={f} value={f}>{f}</option>
                  ))}
                </select>
              ) : (
                <div>
                  <input
                    type="text"
                    value={linkedField}
                    onChange={(e) => setLinkedField(e.target.value.toLowerCase())}
                    required={isLinked}
                    placeholder="e.g. effort"
                    className="w-full max-w-xs px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
                  />
                  <p className="text-xs text-gray-500 mt-1">No linked data uploaded yet. Enter the field name manually (lowercase).</p>
                </div>
              )}
            </div>
          )}
        </div>

        <div className="flex gap-3 justify-end">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
          >
            Cancel
          </button>
          <button
            type="submit"
            disabled={loading}
            className="px-4 py-2 text-sm font-medium text-white bg-primary rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
          >
            {loading ? 'Saving...' : isEditing ? 'Update Group' : 'Create Group'}
          </button>
        </div>
      </form>
    </div>
  );
}
