'use client';

import { useState, useEffect } from 'react';
import { parseComment, countWords } from '@/lib/utils';
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

  const wordCount = countWords(preview);
  const targetWordCount = 100; // Configurable or hardcoded for now
  const percent = Math.min(100, Math.round((wordCount / targetWordCount) * 100));
  const dashArray = 100; 
  const dashOffset = 100 - (percent / 100 * 100); // SVG stroke-dashoffset calculation

  return (
    <div className="flex-1 flex flex-col lg:flex-row gap-8 align-start">
        {/* Sidebar */}
        <aside className="w-full lg:w-80 flex-shrink-0 flex flex-col gap-6">
            <div className="bg-white dark:bg-gray-900 rounded-xl p-6 border border-[#f0f2f4] dark:border-gray-800 shadow-sm">
                <h3 className="text-[#111418] dark:text-white text-sm font-bold uppercase tracking-wider mb-4">Report Categories</h3>
                <div className="flex flex-col gap-2">
                    {groups.map(group => {
                        const selectedOptionId = selections[group.id];
                        const selectedOption = group.options.find(o => o.id === selectedOptionId);
                        const isSelected = !!selectedOption;
                        
                        return (
                            <div key={group.id} className={`flex flex-col gap-2 px-3 py-3 rounded-lg cursor-pointer transition-colors border border-transparent ${isSelected ? 'bg-primary/5 border-primary/10' : 'hover:bg-[#f0f2f4] dark:hover:bg-gray-800'}`}>
                                <div className="flex items-center justify-between">
                                    <div className="flex items-center gap-3">
                                        <span className={`material-symbols-outlined text-xl ${isSelected ? 'text-primary' : 'text-gray-400'}`}>
                                            {group.name === 'Attainment' ? 'school' : 
                                             group.name === 'Effort' ? 'fitness_center' : 
                                             group.name === 'Homework' ? 'home_work' : 'article'}
                                        </span>
                                        <p className={`text-sm font-medium ${isSelected ? 'text-primary' : 'text-[#111418] dark:text-gray-300'}`}>{group.name}</p>
                                    </div>
                                    {isSelected && <span className="bg-primary text-white text-[10px] px-1.5 py-0.5 rounded-full">{selectedOption?.code}</span>}
                                </div>
                                {/* Options (Inline for now or Expandable) - Let's keep them inline for quick access as per logic */}
                                <div className="pl-8 flex flex-wrap gap-2 mt-1">
                                    {group.options.map(opt => (
                                        <button 
                                            key={opt.id}
                                            onClick={() => handleSelection(group.id, opt.id)}
                                            className={`text-xs px-2 py-1 rounded border ${selections[group.id] === opt.id ? 'bg-primary text-white border-primary' : 'bg-white text-gray-600 border-gray-200 hover:border-gray-300'}`}
                                            title={opt.text}
                                        >
                                            {opt.code}
                                        </button>
                                    ))}
                                </div>
                            </div>
                        )
                    })}
                </div>
            </div>

            <div className="bg-white dark:bg-gray-900 rounded-xl p-6 border border-[#f0f2f4] dark:border-gray-800 shadow-sm">
                <h3 className="text-[#111418] dark:text-white text-sm font-bold uppercase tracking-wider mb-4">Selected Codes</h3>
                <div className="flex flex-wrap gap-2">
                     {Object.entries(selections).map(([gid, oid]) => {
                         const group = groups.find(g => g.id === gid);
                         const option = group?.options.find(o => o.id === oid);
                         if (!group || !option) return null;
                         return (
                            <span key={gid} className="bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 px-2 py-1 rounded text-xs font-medium border border-blue-200 dark:border-blue-800">
                                #{group.name}{option.code}
                            </span>
                         )
                     })}
                     {Object.keys(selections).length === 0 && <span className="text-gray-400 text-xs italic">No selection</span>}
                </div>
            </div>
        </aside>

        {/* Editor Section */}
        <div className="flex-1 flex flex-col bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm overflow-hidden min-h-[500px]">
            <div className="px-8 py-5 border-b border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center bg-gray-50/50 dark:bg-gray-800/50">
                <h3 className="text-[#111418] dark:text-white text-lg font-bold leading-tight">Report Preview & Editor</h3>
                <div className="flex items-center gap-4">
                    {/* Progress Ring and Word Count */}
                    <div className="flex items-center gap-3">
                        <div className="relative flex items-center justify-center size-10">
                            <svg className="transform -rotate-90 w-full h-full" viewBox="0 0 40 40">
                                <circle className="text-gray-200 dark:text-gray-700" cx="20" cy="20" fill="transparent" r="16" stroke="currentColor" strokeWidth="3"></circle>
                                <circle 
                                    className="text-primary transition-all duration-500" 
                                    cx="20" cy="20" 
                                    fill="transparent" 
                                    r="16" 
                                    stroke="currentColor" 
                                    strokeDasharray="100" 
                                    strokeDashoffset={100 - percent} 
                                    strokeLinecap="round" 
                                    strokeWidth="3"
                                ></circle>
                            </svg>
                            <span className="absolute text-[10px] font-bold text-primary">{percent}%</span>
                        </div>
                        <div className="flex flex-col">
                            <span className="text-xs font-bold text-[#111418] dark:text-white leading-none">{wordCount}/{targetWordCount} words</span>
                            <span className="text-[10px] text-[#617289] dark:text-gray-400">Target length</span>
                        </div>
                    </div>
                    <div className="h-8 w-[1px] bg-gray-200 dark:bg-gray-700 mx-2"></div>
                    <button 
                        onClick={copyToClipboard}
                        className="text-[#617289] dark:text-gray-400 hover:text-primary transition-colors flex items-center gap-1">
                        <span className="material-symbols-outlined text-lg">{copied ? 'check' : 'content_copy'}</span>
                    </button>
                </div>
            </div>
            
            <div className="flex-1 p-8 relative">
                <div className="relative h-full flex flex-col">
                    <label className="text-xs font-bold text-primary uppercase mb-2 block">Generated Comment</label>
                    <textarea 
                        className="flex-1 w-full p-6 text-lg leading-relaxed text-[#111418] dark:text-gray-100 bg-background-light/50 dark:bg-gray-950/50 rounded-xl border border-transparent focus:border-primary focus:ring-2 focus:ring-primary/20 resize-none outline-none font-display transition-all" 
                        placeholder="Start typing student comments here..."
                        value={preview}
                        onChange={handleTextChange}
                        onBlur={handleTextBlur}
                    ></textarea>
                    
                    {/* Editor Toolbar Floating */}
                    <div className="absolute bottom-6 right-6 flex gap-2 bg-white dark:bg-gray-800 shadow-xl rounded-lg p-1.5 border border-gray-100 dark:border-gray-700">
                        <button className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-600 dark:text-gray-300"><span className="material-symbols-outlined text-lg">format_bold</span></button>
                        <button className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-600 dark:text-gray-300"><span className="material-symbols-outlined text-lg">format_italic</span></button>
                        <button className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-600 dark:text-gray-300"><span className="material-symbols-outlined text-lg">auto_fix_high</span></button>
                        <div className="w-px h-6 bg-gray-200 dark:bg-gray-700 my-auto"></div>
                        <button className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-600 dark:text-gray-300"><span className="material-symbols-outlined text-lg">undo</span></button>
                        <button className="p-2 hover:bg-gray-100 dark:hover:bg-gray-700 rounded text-gray-600 dark:text-gray-300"><span className="material-symbols-outlined text-lg">redo</span></button>
                    </div>
                </div>
            </div>
            
            <div className="px-8 py-4 bg-primary/5 dark:bg-primary/10 border-t border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center">
                <div className="flex items-center gap-2 text-primary">
                    <span className="material-symbols-outlined text-base">lightbulb</span>
                    <span className="text-xs font-semibold">Tip: Changes are saved automatically.</span>
                </div>

            </div>
        </div>
    </div>
  );
}
