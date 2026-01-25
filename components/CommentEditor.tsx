'use client';

import { useState, useEffect } from 'react';
import { parseComment, countWords } from '@/lib/utils';
import { Copy, Check } from 'lucide-react';
import { updateAssignmentCode, updateAssignmentCommentText } from '@/app/actions';

type CommentOption = {
  id: string;
  code: string;
  text: string;
};

type CommentGroup = {
  id: string;
  name: string;
  options: CommentOption[];
};

type Pupil = {
  admissionNumber: string;
  firstName: string;
  lastName: string;
  gender: string;
};

type PupilCode = {
  groupId: string;
  code: string | null;
};

type Assignment = {
  id: string;
  pupil: Pupil;
  codes: PupilCode[];
  finalComment?: string | null;
  eoyLevel?: string | null;
  targetLevel?: string | null;
  class?: {
    name: string;
    year?: string | null;
  };
};

type Subject = {
  id: string;
  title: string | null;
  studiedComment?: string | null;
};

interface CommentEditorProps {
  assignment: Assignment;
  subject: Subject;
  groups: CommentGroup[];
}

export default function CommentEditor({ assignment, subject, groups }: CommentEditorProps) {
  const [selections, setSelections] = useState<Record<string, string>>({}); // groupId -> optionId
  const [preview, setPreview] = useState('');
  const [copied, setCopied] = useState(false);

  // Initialize selections with assignment's pre-assigned codes
  useEffect(() => {
    const initialSelections: Record<string, string> = {};
    
    assignment.codes.forEach(pc => {
        const group = groups.find(g => g.id === pc.groupId);
        if (group && pc.code) {
            const option = group.options.find(o => o.code === pc.code);
            if (option) {
                initialSelections[group.id] = option.id;
            }
        }
    });

    setSelections(initialSelections);
  }, [assignment, groups]);

  const [initialLoad, setInitialLoad] = useState(true);

  // Generate Preview
  useEffect(() => {
    if (initialLoad) {
        if (assignment.finalComment) {
            setPreview(assignment.finalComment);
            setInitialLoad(false);
            return;
        }
        setInitialLoad(false);
    }
    
    const parts: string[] = [];

    // 1. Studied Comment (Subject Intro)
    if (subject.studiedComment) {
      parts.push(subject.studiedComment);
    }

    // 2. Groups ordered by display order
    const sortedGroups = [...groups].sort((a,b) => (a as any).displayOrder || 0 - ((b as any).displayOrder || 0));
    
    // Formatting logic similar to legacy
    // Let's identify WP, TH, PS, OA for legacy layout if they exist
    const wp = sortedGroups.find(g => g.name === 'WP');
    const th = sortedGroups.find(g => g.name === 'TH');
    const ps = sortedGroups.find(g => g.name === 'PS');
    const oa = sortedGroups.find(g => g.name === 'OA');
    const others = sortedGroups.filter(g => !['WP', 'TH', 'PS', 'OA'].includes(g.name));

    const getOptText = (group?: CommentGroup) => {
        if (!group) return "";
        const selectedOptionId = selections[group.id];
        if (!selectedOptionId) return "";
        return group.options.find(o => o.id === selectedOptionId)?.text || "";
    };

    const wpText = getOptText(wp);
    const thText = getOptText(th);
    const psText = getOptText(ps);
    const oaText = getOptText(oa);

    let combined = "";
    if (subject.studiedComment) combined += subject.studiedComment + "\n\n";
    
    const middleBlock = [wpText, thText, psText].filter(Boolean).join(" ");
    if (middleBlock) combined += middleBlock + "\n\n";

    if (oaText) combined += oaText;

    // Handle unknown groups
    others.forEach(g => {
        const t = getOptText(g);
        if (t) combined += "\n\n" + t;
    });

    setPreview(parseComment(
      combined, 
      assignment.pupil.firstName, 
      assignment.pupil.gender,
      subject.title || '',
      assignment.class?.year,
      assignment.eoyLevel,
      assignment.targetLevel
    ));

  }, [selections, subject, groups, assignment, initialLoad]);

  const handleSelection = async (groupId: string, optionId: string) => {
    setSelections(prev => ({
      ...prev,
      [groupId]: optionId
    }));
    
    const group = groups.find(g => g.id === groupId);
    const option = group?.options.find(o => o.id === optionId);
    
    if (option) {
        await updateAssignmentCode(assignment.id, groupId, option.code);
    }
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(preview);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setPreview(e.target.value);
  };

  const handleTextBlur = async () => {
    await updateAssignmentCommentText(assignment.id, preview);
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 h-[calc(100vh-140px)]">
        <div className="overflow-y-auto pr-4 space-y-8">
            {groups.map(group => (
                <div key={group.id} className="bg-white p-6 rounded-lg shadow-sm border border-gray-100">
                    <h3 className="text-lg font-semibold text-gray-800 mb-4">{group.name} Comments</h3>
                    <div className="space-y-3">
                        {group.options.map(option => (
                           <label key={option.id} className="flex items-start gap-3 p-3 rounded-md hover:bg-gray-50 cursor-pointer border border-transparent hover:border-gray-200 transition-all select-none">
                              <input 
                                type="radio" 
                                name={group.id} 
                                value={option.id}
                                checked={selections[group.id] === option.id}
                                onChange={() => handleSelection(group.id, option.id)}
                                className="mt-1 w-4 h-4 text-blue-600 bg-gray-100 border-gray-300 focus:ring-blue-500"
                              />
                              <div className="flex-1">
                                <div className="flex items-center justify-between mb-1">
                                    <span className="text-xs font-bold text-gray-500 uppercase tracking-wide">{option.code}</span>
                                </div>
                                <p className="text-sm text-gray-700">{parseComment(
                                  option.text, 
                                  assignment.pupil.firstName, 
                                  assignment.pupil.gender,
                                  subject.title || '',
                                  assignment.class?.year,
                                  assignment.eoyLevel,
                                  assignment.targetLevel
                                )}</p>
                              </div>
                           </label>
                        ))}
                    </div>
                </div>
            ))}
        </div>

        <div className="bg-white rounded-xl shadow-lg border border-gray-200 flex flex-col h-full sticky top-8">
            <div className="p-4 border-b border-gray-100 bg-gray-50 flex justify-between items-center rounded-t-xl">
                <h2 className="font-semibold text-gray-700">Preview (Editable)</h2>
                <div className="flex gap-2">
                    <button 
                        onClick={copyToClipboard}
                        className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-md text-sm font-medium transition-colors"
                    >
                        {copied ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                        {copied ? 'Copied!' : 'Copy Text'}
                    </button>
                </div>
            </div>
            <div className="flex-1 p-0 overflow-hidden flex flex-col">
                <textarea
                    value={preview}
                    onChange={handleTextChange}
                    onBlur={handleTextBlur}
                    className="flex-1 w-full p-6 text-lg text-gray-800 leading-relaxed border-none resize-none focus:ring-0 focus:outline-none"
                    placeholder="Select comments to generate a preview..."
                />
            </div>
            <div className="px-4 py-2 bg-gray-50 text-xs text-gray-400 border-t items-center flex justify-between">
                <span>Changes saved automatically on blur.</span>
                <span className="font-medium">{countWords(preview)} words</span>
            </div>
        </div>
    </div>
  );
}
