'use client';

import { useState, useEffect } from 'react';
import { parseComment } from '@/lib/utils';
import { Copy, Check } from 'lucide-react';
import { updateStudentComment, updateStudentCommentText } from '@/app/actions';

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

type Student = {
  id: string;
  firstName: string;
  lastName: string;
  gender: string;
  wpCode?: string | null;
  thCode?: string | null;
  psCode?: string | null;
  oaCode?: string | null;
  comment?: string | null;
};

type Course = {
  id: string;
  name: string;
  studiedComment?: string | null;
};

interface CommentEditorProps {
  student: Student;
  course: Course;
  groups: CommentGroup[];
}

export default function CommentEditor({ student, course, groups }: CommentEditorProps) {
  const [selections, setSelections] = useState<Record<string, string>>({});
  const [preview, setPreview] = useState('');
  const [copied, setCopied] = useState(false);

  // Initialize selections with student's pre-assigned codes if available
  useEffect(() => {
    const initialSelections: Record<string, string> = {};
    
    // Helper to find option ID by code
    const findOptionId = (groupName: string, code: string | null | undefined) => {
      if (!code) return undefined;
      const group = groups.find(g => g.name === groupName);
      const option = group?.options.find(o => o.code === code);
      return option?.id;
    };

    const wpId = findOptionId('WP', student.wpCode);
    if (wpId) initialSelections['WP'] = wpId;

    const thId = findOptionId('TH', student.thCode);
    if (thId) initialSelections['TH'] = thId;

    const psId = findOptionId('PS', student.psCode);
    if (psId) initialSelections['PS'] = psId;

    const oaId = findOptionId('OA', student.oaCode);
    if (oaId) initialSelections['OA'] = oaId;

    setSelections(initialSelections);
  }, [student, groups]);

  // Track if we have done the initial load
  const [initialLoad, setInitialLoad] = useState(true);

  // Generate Preview
  useEffect(() => {
    // If it's the very first load and we have a saved comment, use that.
    // Discard the generated text this one time.
    if (initialLoad) {
        if (student.comment) {
            setPreview(student.comment);
            setInitialLoad(false);
            return;
        }
        setInitialLoad(false);
    }
    
    // Normal generation logic follows...
    const parts: string[] = [];

    // 1. Studied Comment (Course Intro)
    if (course.studiedComment) {
      parts.push(course.studiedComment);
    }

    // 2. Groups in specific order: WP, TH, PS, OA
    const order = ['WP', 'TH', 'PS', 'OA'];
    // And others?
    const otherGroups = groups.filter(g => !order.includes(g.name)).map(g => g.name);
    const finalOrder = [...order, ...otherGroups];

    finalOrder.forEach(groupName => {
        // Find group ID from name (we stored selections by Group NAME for easier ordering?)
        // Actually selections key is group NAME for simplicity in this logic
        const selectedOptionId = selections[groupName];
        if (selectedOptionId) {
            const group = groups.find(g => g.name === groupName);
            const option = group?.options.find(o => o.id === selectedOptionId);
            if (option) {
                parts.push(option.text);
            }
        }
    });

    const rawText = parts.join('\n\n'); // Separate paragraphs? Or spaces? 
    // Legacy app used: studiedComment + " " + wp + " " + th + " " + ps + "\n" + oa
    // logic:
    // studiedComment
    // wpComment + " " + thComment + " " + psComment
    // oaComment
    
    // Let's replicate this specific formatting for known groups.
    const studied = course.studiedComment || "";
    
    const getOptText = (name: string) => {
        const id = selections[name];
        if (!id) return "";
        const g = groups.find(g => g.name === name);
        return g?.options.find(o => o.id === id)?.text || "";
    };

    const wp = getOptText("WP");
    const th = getOptText("TH");
    const ps = getOptText("PS");
    const oa = getOptText("OA");

    let combined = "";
    if (studied) combined += studied + "\n\n";
    
    const middleBlock = [wp, th, ps].filter(Boolean).join(" ");
    if (middleBlock) combined += middleBlock + "\n\n";

    if (oa) combined += oa;

    // Handle unknown groups (append at end)
    otherGroups.forEach(name => {
        const t = getOptText(name);
        if (t) combined += "\n\n" + t;
    });

    setPreview(parseComment(combined, student.firstName, student.gender));

  }, [selections, course, groups, student]);

  const handleSelection = async (groupName: string, optionId: string) => {
    // Optimistic update
    setSelections(prev => ({
      ...prev,
      [groupName]: optionId
    }));
    
    // Find the code for this option
    const group = groups.find(g => g.name === groupName);
    const option = group?.options.find(o => o.id === optionId);
    
    if (option) {
        await updateStudentComment(student.id, groupName, option.code);
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
    // Persist on blur
    await updateStudentCommentText(student.id, preview);
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 h-[calc(100vh-140px)]">
        {/* Left: Controls */}
        <div className="overflow-y-auto pr-4 space-y-8">
            {groups.sort((a,b) => {
                const order = ['WP', 'TH', 'PS', 'OA'];
                 const indexA = order.indexOf(a.name);
                 const indexB = order.indexOf(b.name);
                 
                 // If both are in the order list, sort by index
                 if (indexA !== -1 && indexB !== -1) return indexA - indexB;
                 
                 // If only A is in list, A comes first
                 if (indexA !== -1) return -1;
                 
                 // If only B is in list, B comes first
                 if (indexB !== -1) return 1;
                 
                 // Otherwise sort alphabetically
                 return a.name.localeCompare(b.name);
            }).map(group => (
                <div key={group.id} className="bg-white p-6 rounded-lg shadow-sm border border-gray-100">
                    <h3 className="text-lg font-semibold text-gray-800 mb-4">{group.name} Comments</h3>
                    <div className="space-y-3">
                        {group.options.map(option => (
                           <label key={option.id} className="flex items-start gap-3 p-3 rounded-md hover:bg-gray-50 cursor-pointer border border-transparent hover:border-gray-200 transition-all select-none">
                              <input 
                                type="radio" 
                                name={group.name} 
                                value={option.id}
                                checked={selections[group.name] === option.id}
                                onChange={() => handleSelection(group.name, option.id)}
                                className="mt-1 w-4 h-4 text-blue-600 bg-gray-100 border-gray-300 focus:ring-blue-500"
                              />
                              <div className="flex-1">
                                <div className="flex items-center justify-between mb-1">
                                    <span className="text-xs font-bold text-gray-500 uppercase tracking-wide">{option.code}</span>
                                </div>
                                <p className="text-sm text-gray-700">{parseComment(option.text, student.firstName, student.gender)}</p>
                              </div>
                           </label>
                        ))}
                    </div>
                </div>
            ))}
        </div>

        {/* Right: Preview (Editable) */}
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
            </div>
        </div>
    </div>
  );
}
