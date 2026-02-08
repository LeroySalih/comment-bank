'use client';

import { useState } from 'react';
import { useRouter } from 'next/navigation';
import { updateWrapperTemplate } from '@/lib/server-actions/ccg';

type Group = {
  id: string;
  name: string;
};

interface WrapperTemplateEditorProps {
  template: string;
  p2Groups: Group[];
}

export default function WrapperTemplateEditor({ template, p2Groups }: WrapperTemplateEditorProps) {
  const router = useRouter();
  const [value, setValue] = useState(template);
  const [loading, setLoading] = useState(false);
  const [saved, setSaved] = useState(false);

  const handleSave = async () => {
    setLoading(true);
    try {
      const result = await updateWrapperTemplate(value);
      if (result.success) {
        setSaved(true);
        setTimeout(() => setSaved(false), 2000);
        router.refresh();
      } else {
        alert((result as any).error || 'Failed to save template');
      }
    } catch {
      alert('Failed to save template');
    }
    setLoading(false);
  };

  return (
    <div className="bg-white border border-gray-200 rounded-xl shadow-sm p-6 space-y-4">
      <div>
        <h3 className="text-base font-bold text-gray-900">P2 Wrapper Template</h3>
        <p className="text-sm text-gray-500 mt-1">
          Controls how Paragraph 2 groups are combined. Use placeholders like{' '}
          {p2Groups.map((g, i) => (
            <span key={g.id}>
              {i > 0 && ', '}
              <code className="bg-gray-100 px-1 rounded text-xs">{`<${g.name}>`}</code>
            </span>
          ))}
        </p>
      </div>

      <textarea
        value={value}
        onChange={(e) => setValue(e.target.value)}
        rows={3}
        className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-primary focus:border-transparent resize-none font-mono"
        placeholder="e.g. <Effort> <Behaviour> <Homework>"
      />

      <div className="flex items-center justify-between">
        <div className="flex flex-wrap gap-2">
          {p2Groups.map(g => (
            <span key={g.id} className="text-xs bg-blue-50 text-blue-700 px-2 py-1 rounded font-mono">
              {`<${g.name}>`}
            </span>
          ))}
        </div>
        <button
          onClick={handleSave}
          disabled={loading || value === template}
          className="px-4 py-2 text-sm font-medium text-white bg-primary rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
        >
          {loading ? 'Saving...' : saved ? 'Saved!' : 'Save Template'}
        </button>
      </div>
    </div>
  );
}
