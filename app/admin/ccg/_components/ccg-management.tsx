'use client';

import { useState } from 'react';
import CcgGroupCard from './ccg-group-card';
import CcgGroupForm from './ccg-group-form';
import WrapperTemplateEditor from './wrapper-template-editor';

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

interface CcgManagementProps {
  groups: Group[];
  wrapperTemplate: string;
  availableLinkedFields?: string[];
}

const POSITION_LABELS: Record<string, string> = {
  p1: 'Paragraph 1 — Introduction',
  p2: 'Paragraph 2 — General Attributes',
  p4: 'Paragraph 4 — Conclusion',
};

const POSITION_DESCRIPTIONS: Record<string, string> = {
  p1: 'Opening statements about academic performance',
  p2: 'Combined using the wrapper template below',
  p4: 'Closing/summary statements',
};

export default function CcgManagement({ groups, wrapperTemplate, availableLinkedFields = [] }: CcgManagementProps) {
  const [showCreateForm, setShowCreateForm] = useState(false);

  const groupsByPosition: Record<string, Group[]> = { p1: [], p2: [], p4: [] };
  groups.forEach(g => {
    if (groupsByPosition[g.paragraphPosition]) {
      groupsByPosition[g.paragraphPosition].push(g);
    }
  });

  return (
    <div className="space-y-8">
      {/* Create Group Button */}
      <div className="flex justify-end">
        <button
          onClick={() => setShowCreateForm(true)}
          className="flex items-center gap-2 px-4 py-2 bg-primary text-white rounded-lg hover:bg-blue-700 transition-colors font-medium text-sm"
        >
          <span className="material-symbols-outlined text-lg">add</span>
          Add Group
        </button>
      </div>

      {showCreateForm && (
        <CcgGroupForm availableLinkedFields={availableLinkedFields} onClose={() => setShowCreateForm(false)} />
      )}

      {/* Groups organized by paragraph position */}
      {(['p1', 'p2', 'p4'] as const).map(position => (
        <div key={position} className="space-y-4">
          <div className="border-b border-gray-200 pb-2">
            <h2 className="text-xl font-bold text-gray-900">{POSITION_LABELS[position]}</h2>
            <p className="text-sm text-gray-500">{POSITION_DESCRIPTIONS[position]}</p>
          </div>

          {/* Note about P3 after P2 */}
          {position === 'p2' && (
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
              <p className="text-sm text-blue-700">
                <span className="font-semibold">Paragraph 3</span> is automatically generated from each subject&apos;s own comment groups. It is not managed here.
              </p>
            </div>
          )}

          {groupsByPosition[position].length === 0 ? (
            <div className="bg-gray-50 border border-gray-200 rounded-lg p-8 text-center">
              <p className="text-gray-500 text-sm">No groups in this position yet</p>
            </div>
          ) : (
            <div className="space-y-4">
              {groupsByPosition[position].map(group => (
                <CcgGroupCard key={group.id} group={group} availableLinkedFields={availableLinkedFields} />
              ))}
            </div>
          )}

          {/* Wrapper template editor for P2 */}
          {position === 'p2' && (
            <WrapperTemplateEditor
              template={wrapperTemplate}
              p2Groups={groupsByPosition.p2}
            />
          )}
        </div>
      ))}
    </div>
  );
}
