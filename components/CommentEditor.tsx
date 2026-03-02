'use client';

import { useState, useEffect, useRef } from 'react';
import { useRouter } from 'next/navigation';
import { parseComment, countWords } from '@/lib/utils';
import { updateAssignmentCode, updateCommonAssignmentCode, updateAssignmentCommentText, revertAssignmentComment } from '@/app/actions';
import { reviewComment } from '@/lib/server-actions/comment-check';
import CommentStatusBadge from './CommentStatusBadge';
import ConfirmModal from './ConfirmModal';

type CommentOption = {
  id: string;
  code: string;
  text: string;
};

type CommentGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommentOption: CommentOption[];
};

type CommonCommentOption = {
  id: string;
  code: string;
  text: string;
};

type CommonCommentGroup = {
  id: string;
  name: string;
  title: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommonCommentOption: CommonCommentOption[];
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

type CommonPupilCode = {
  commonGroupId: string;
  code: string | null;
};

type Assignment = {
  id: string;
  Pupil: Pupil;
  PupilCode: PupilCode[];
  CommonPupilCode?: CommonPupilCode[];
  finalComment?: string | null;
  eoyLevel?: string | null;
  targetLevel?: string | null;
  checkStatus?: string;
  checkNote?: string | null;
  linkedData?: any;
  Class?: {
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
  isHoD?: boolean;
  commonGroups?: CommonCommentGroup[];
  formatTemplate?: string;
  subjectFormat?: string | null;
}

export default function CommentEditor({ assignment, subject, groups, isHoD = false, commonGroups, formatTemplate, subjectFormat }: CommentEditorProps) {
  const router = useRouter();
  const [selections, setSelections] = useState<Record<string, string>>({}); // groupId -> optionId
  const [commonSelections, setCommonSelections] = useState<Record<string, string>>({}); // commonGroupId -> optionId
  const [preview, setPreview] = useState('');
  const [copied, setCopied] = useState(false);
  const [checkStatus, setCheckStatus] = useState(assignment.checkStatus || 'not_required');
  const [isManuallyEdited, setIsManuallyEdited] = useState(false);
  const [showRevertModal, setShowRevertModal] = useState(false);
  const [isReverting, setIsReverting] = useState(false);

  // HoD review state
  const [rejectionNote, setRejectionNote] = useState('');
  const [isReviewing, setIsReviewing] = useState(false);

  // Initialize selections with assignment's pre-assigned codes
  useEffect(() => {
    const initialSelections: Record<string, string> = {};
    const linkedData = assignment.linkedData as Record<string, string> | null | undefined;

    // Subject-specific groups
    for (const group of groups) {
      if (group.isLinked && group.linkedField && linkedData) {
        const code = linkedData[group.linkedField];
        if (code) {
          const option = group.CommentOption.find(o => o.code === code);
          if (option) initialSelections[group.id] = option.id;
        }
      } else {
        const pc = assignment.PupilCode.find(p => p.groupId === group.id);
        if (pc?.code) {
          const option = group.CommentOption.find(o => o.code === pc.code);
          if (option) initialSelections[group.id] = option.id;
        }
      }
    }
    setSelections(initialSelections);

    // Initialize common selections
    if (commonGroups) {
      const initialCommon: Record<string, string> = {};
      for (const cg of commonGroups) {
        if (cg.isLinked && cg.linkedField && linkedData) {
          const code = linkedData[cg.linkedField];
          if (code) {
            const option = cg.CommonCommentOption.find(o => o.code === code);
            if (option) initialCommon[cg.id] = option.id;
          }
        } else {
          const cpc = (assignment.CommonPupilCode || []).find(c => c.commonGroupId === cg.id);
          if (cpc?.code) {
            const option = cg.CommonCommentOption.find(o => o.code === cpc.code);
            if (option) initialCommon[cg.id] = option.id;
          }
        }
      }
      setCommonSelections(initialCommon);
    }
  }, [assignment, groups, commonGroups]);

  // Track if we've loaded the initial comment and should skip regeneration
  const hasLoadedInitialComment = useRef(false);
  const skipRegeneration = useRef(false);

  // Initialize preview from finalComment when it becomes available
  useEffect(() => {
    if (!hasLoadedInitialComment.current && assignment.finalComment) {
      setPreview(assignment.finalComment);
      hasLoadedInitialComment.current = true;
      skipRegeneration.current = true;
    }
  }, [assignment.finalComment]);

  // Generate Preview from selections (only when selections change and not manually edited)
  useEffect(() => {
    if (isManuallyEdited) return;
    if (skipRegeneration.current) return;
    if (!hasLoadedInitialComment.current && assignment.finalComment) return;

    // Helper to get option text for subject-specific groups
    const getOptText = (group: CommentGroup) => {
        const selectedOptionId = selections[group.id];
        if (!selectedOptionId) return "";
        return group.CommentOption.find(o => o.id === selectedOptionId)?.text || "";
    };

    // Helper to get option text for common groups
    const getCommonOptText = (group: CommonCommentGroup) => {
        const selectedOptionId = commonSelections[group.id];
        if (!selectedOptionId) return "";
        return group.CommonCommentOption.find(o => o.id === selectedOptionId)?.text || "";
    };

    // Identify subject groups that override CCG variables
    const ccgGroupNames = new Set((commonGroups ?? []).map(g => g.name));
    const overrideGroupsByName = new Map(
        groups
            .filter(g => ccgGroupNames.has(g.name))
            .map(g => [g.name, g] as [string, typeof groups[number]])
    );

    // Build the subject content
    const buildSubjectContent = (): string => {
        const parts: string[] = [];
        if (subject.studiedComment) parts.push(subject.studiedComment);

        // Exclude CCG override groups from SCG content
        const pureScgGroups = groups.filter(g => !ccgGroupNames.has(g.name));

        let orderedGroups: CommentGroup[];
        if (subjectFormat) {
            const codes = subjectFormat.split(/\s+/).filter(Boolean);
            orderedGroups = [];
            for (const code of codes) {
                const group = pureScgGroups.find(g => g.name === code);
                if (group) orderedGroups.push(group);
            }
        } else {
            orderedGroups = [...pureScgGroups].sort((a, b) => ((a as any).displayOrder || 0) - ((b as any).displayOrder || 0));
        }

        const groupTexts: string[] = [];
        for (const group of orderedGroups) {
            const text = getOptText(group);
            if (text) groupTexts.push(text);
        }
        if (groupTexts.length > 0) parts.push(groupTexts.join(" "));
        return parts.join(" ");
    };

    if (formatTemplate && commonGroups && commonGroups.length > 0) {
      // Format template-based generation
      let result = formatTemplate;

      // Replace each CCG group tag — use subject override if available, fall back to CCG
      for (const group of commonGroups) {
        const overrideGroup = overrideGroupsByName.get(group.name);
        let text: string;
        if (overrideGroup) {
            text = getOptText(overrideGroup) || getCommonOptText(group);
        } else {
            text = getCommonOptText(group);
        }
        result = result.replaceAll(`<${group.name}>`, text);
      }

      // Replace <SCG> tag with subject comment group content
      const subjectContent = buildSubjectContent();
      result = result.replaceAll('<SCG>', subjectContent);

      // Clean up unreplaced custom tags (not standard variables)
      const standardVars = ['Name', 'He', 'he', 'She', 'she', 'His', 'his', 'Her', 'her', 'Him', 'him', 'Subject', 'TargetLevel', 'EoYLevel', 'Year', 'SCG'];
      result = result.replace(/<([^>]+)>/g, (match, tagName) => {
        if (standardVars.includes(tagName)) return match;
        return '';
      });
      result = result.split(/\n+/).map(p => p.replace(/\s+/g, ' ').trim()).filter(p => p).join('\n\n');

      setPreview(parseComment(result, assignment.Pupil.firstName, assignment.Pupil.gender, subject.title || '', assignment.Class?.year, assignment.eoyLevel, assignment.targetLevel));
    } else if (!commonGroups || commonGroups.length === 0) {
      // No common groups — just subject content
      const subjectContent = buildSubjectContent();
      setPreview(parseComment(subjectContent, assignment.Pupil.firstName, assignment.Pupil.gender, subject.title || '', assignment.Class?.year, assignment.eoyLevel, assignment.targetLevel));
    } else {
      // Has common groups but no template: join common then subject
      const paragraphs: string[] = [];
      const commonTexts = commonGroups.map(g => {
          const overrideGroup = overrideGroupsByName.get(g.name);
          if (overrideGroup) {
              return getOptText(overrideGroup) || getCommonOptText(g);
          }
          return getCommonOptText(g);
      }).filter(Boolean);
      if (commonTexts.length > 0) paragraphs.push(commonTexts.join(" "));

      const subjectContent = buildSubjectContent();
      if (subjectContent) paragraphs.push(subjectContent);

      const combined = paragraphs.join("\n\n");
      setPreview(parseComment(combined, assignment.Pupil.firstName, assignment.Pupil.gender, subject.title || '', assignment.Class?.year, assignment.eoyLevel, assignment.targetLevel));
    }

  }, [selections, commonSelections, subject, groups, assignment, isManuallyEdited, commonGroups, formatTemplate, subjectFormat]);

  const handleSelection = async (groupId: string, optionId: string) => {
    setSelections(prev => ({
      ...prev,
      [groupId]: optionId
    }));

    const group = groups.find(g => g.id === groupId);
    const option = group?.CommentOption.find(o => o.id === optionId);

    if (option) {
        await updateAssignmentCode(assignment.id, groupId, option.code);
    }
  };

  const handleCommonSelection = async (commonGroupId: string, optionId: string) => {
    setCommonSelections(prev => ({
      ...prev,
      [commonGroupId]: optionId
    }));

    const group = commonGroups?.find(g => g.id === commonGroupId);
    const option = group?.CommonCommentOption.find(o => o.id === optionId);

    if (option) {
      await updateCommonAssignmentCode(assignment.id, commonGroupId, option.code);
    }
  };

  const copyToClipboard = () => {
    navigator.clipboard.writeText(preview);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setPreview(e.target.value);
    setIsManuallyEdited(true);
  };

  const handleTextBlur = async () => {
    if (!isManuallyEdited) return;

    const previousStatus = checkStatus;
    setCheckStatus('required_check');

    try {
      const result = await updateAssignmentCommentText(assignment.id, preview);
      if (!result.success) {
        console.error('Failed to save comment:', result.error);
        setCheckStatus(previousStatus);
        alert('Failed to save comment: ' + (result.error || 'Unknown error'));
      }
    } catch (error) {
      console.error('Error saving comment:', error);
      setCheckStatus(previousStatus);
      alert('Error saving comment. Check console for details.');
    }
  };

  const handleRevert = async () => {
    setIsReverting(true);
    try {
      const result = await revertAssignmentComment(assignment.id);
      if (result.success) {
        setIsManuallyEdited(false);
        skipRegeneration.current = false;
        hasLoadedInitialComment.current = false;
        setCheckStatus('not_required');
        setShowRevertModal(false);
        router.refresh();
      } else {
        alert('Failed to revert: ' + result.error);
      }
    } catch (error) {
      alert('Failed to revert comment');
    }
    setIsReverting(false);
  };

  const handleApprove = async () => {
    setIsReviewing(true);
    try {
      const result = await reviewComment(assignment.id, 'checked_ok');
      if (result.success) {
        setCheckStatus('checked_ok');
        router.refresh();
      } else {
        alert('Failed to approve: ' + result.error);
      }
    } catch (error) {
      alert('Failed to approve comment');
    }
    setIsReviewing(false);
  };

  const handleReject = async () => {
    if (!rejectionNote.trim()) {
      alert('Please provide a reason for rejection');
      return;
    }
    setIsReviewing(true);
    try {
      const result = await reviewComment(assignment.id, 'checked_rejected', rejectionNote);
      if (result.success) {
        setCheckStatus('checked_rejected');
        setRejectionNote('');
        router.refresh();
      } else {
        alert('Failed to reject: ' + result.error);
      }
    } catch (error) {
      alert('Failed to reject comment');
    }
    setIsReviewing(false);
  };

  const wordCount = countWords(preview);
  const targetWordCount = 100;
  const percent = Math.min(100, Math.round((wordCount / targetWordCount) * 100));

  const commentBanksDisabled = isManuallyEdited || skipRegeneration.current;

  // Check linked data mismatch for a group
  const getLinkedMismatch = (groupId: string, linkedField: string | null | undefined, options: { code: string }[]): string | null => {
    if (!linkedField) return null;
    const linkedData = assignment.linkedData as Record<string, string> | null | undefined;
    const code = linkedData?.[linkedField];
    if (!code) return `No data for "${linkedField}"`;
    if (!options.some(o => o.code === code)) return `Value "${code}" has no matching option`;
    return null;
  };

  // Build ordered sections for sidebar
  const renderGroupSection = (
    sectionLabel: string,
    sectionGroups: { id: string; name: string; isLinked?: boolean; linkedField?: string | null; options: { id: string; code: string; text: string }[] }[],
    selectionsMap: Record<string, string>,
    onSelect: (groupId: string, optionId: string) => void,
    isCommon: boolean
  ) => {
    if (sectionGroups.length === 0) return null;
    return (
      <div className="space-y-1">
        <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest px-3">{sectionLabel}</p>
        {sectionGroups.map(group => {
          const selectedOptionId = selectionsMap[group.id];
          const selectedOption = group.options.find(o => o.id === selectedOptionId);
          const isSelected = !!selectedOption;
          const isGroupLinked = !!group.isLinked;
          const linkedMismatch = isGroupLinked ? getLinkedMismatch(group.id, group.linkedField, group.options) : null;
          const isGroupDisabled = commentBanksDisabled || isGroupLinked;

          return (
            <div key={group.id} className={`flex flex-col gap-2 px-3 py-3 rounded-lg transition-colors border border-transparent ${isGroupLinked ? 'bg-purple-50/50 border-purple-100' : ''} ${isGroupDisabled ? 'cursor-not-allowed' : 'cursor-pointer'} ${isSelected && !isGroupLinked ? 'bg-primary/5 border-primary/10' : isGroupDisabled ? '' : 'hover:bg-[#f0f2f4] dark:hover:bg-gray-800'}`}>
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <span className={`material-symbols-outlined text-xl ${isGroupLinked ? 'text-purple-500' : isSelected ? 'text-primary' : 'text-gray-400'}`}>
                    {isGroupLinked ? 'link' : isCommon ? 'public' : 'article'}
                  </span>
                  <p className={`text-sm font-medium ${isGroupLinked ? 'text-purple-700' : isSelected ? 'text-primary' : 'text-[#111418] dark:text-gray-300'}`}>{group.name}</p>
                  {isGroupLinked && (
                    <span className="text-[10px] font-medium text-purple-600 bg-purple-100 px-1.5 py-0.5 rounded">Linked</span>
                  )}
                </div>
                {isSelected && <span className={`text-white text-[10px] px-1.5 py-0.5 rounded-full ${isGroupLinked ? 'bg-purple-500' : 'bg-primary'}`}>{selectedOption?.code}</span>}
              </div>
              {linkedMismatch && (
                <div className="pl-8 flex items-center gap-1 text-amber-600">
                  <span className="material-symbols-outlined text-sm">warning</span>
                  <span className="text-xs">{linkedMismatch}</span>
                </div>
              )}
              <div className="pl-8 flex flex-wrap gap-2 mt-1">
                {group.options.map(opt => (
                  <button
                    key={opt.id}
                    onClick={() => !isGroupDisabled && onSelect(group.id, opt.id)}
                    disabled={isGroupDisabled}
                    className={`text-xs px-2 py-1 rounded border ${selectionsMap[group.id] === opt.id ? (isGroupLinked ? 'bg-purple-500 text-white border-purple-500' : 'bg-primary text-white border-primary') : 'bg-white text-gray-600 border-gray-200'} ${isGroupDisabled ? 'cursor-not-allowed opacity-50' : 'hover:border-gray-300'}`}
                    title={isGroupLinked ? 'Linked — auto-populated from data' : commentBanksDisabled ? 'Comment banks are locked' : opt.text}
                  >
                    {opt.code}
                  </button>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  // Build sections — ordered by template position
  const commonGroupsMapped = (commonGroups || []).map(g => ({ id: g.id, name: g.name, isLinked: g.isLinked, linkedField: g.linkedField, options: g.CommonCommentOption }));
  const subjectGroupsMapped = groups.map(g => ({ id: g.id, name: g.name, isLinked: g.isLinked, linkedField: g.linkedField, options: g.CommentOption }));

  // Split subject groups into CCG overrides and pure SCG groups
  const ccgGroupNames = new Set((commonGroups || []).map(g => g.name));
  const overrideGroupsMapped = subjectGroupsMapped.filter(g => ccgGroupNames.has(g.name));
  const pureScgGroupsMapped = subjectGroupsMapped.filter(g => !ccgGroupNames.has(g.name));

  // Split CCG groups into before-SCG and after-SCG based on where <SCG> appears in the template
  let ccgBeforeSCG = commonGroupsMapped;
  let ccgAfterSCG: typeof commonGroupsMapped = [];

  if (formatTemplate && formatTemplate.includes('<SCG>')) {
    const scgIdx = formatTemplate.indexOf('<SCG>');
    const namesAfterSCG = new Set([...formatTemplate.slice(scgIdx + 5).matchAll(/<([^>]+)>/g)].map(m => m[1]));
    const posMap = new Map([...formatTemplate.matchAll(/<([^>]+)>/g)].map((m, i) => [m[1], i] as [string, number]));
    const byPos = (a: { name: string }, b: { name: string }) => (posMap.get(a.name) ?? 999) - (posMap.get(b.name) ?? 999);
    ccgBeforeSCG = commonGroupsMapped.filter(g => !namesAfterSCG.has(g.name)).sort(byPos);
    ccgAfterSCG = commonGroupsMapped.filter(g => namesAfterSCG.has(g.name)).sort(byPos);
  }

  return (
    <div className="flex-1 flex flex-col lg:flex-row gap-8 align-start">
        {/* Sidebar */}
        <aside className="w-full lg:w-80 flex-shrink-0 flex flex-col gap-6">
            <div className={`bg-white dark:bg-gray-900 rounded-xl p-6 border border-[#f0f2f4] dark:border-gray-800 shadow-sm ${commentBanksDisabled ? 'opacity-60' : ''}`}>
                <div className="flex items-center justify-between mb-4">
                    <h3 className="text-[#111418] dark:text-white text-sm font-bold uppercase tracking-wider">Report Categories</h3>
                    {commentBanksDisabled && (
                        <span className="text-xs text-amber-600 dark:text-amber-400 flex items-center gap-1">
                            <span className="material-symbols-outlined text-sm">lock</span>
                            Locked
                        </span>
                    )}
                </div>
                {commentBanksDisabled && (
                    <div className="mb-4 p-3 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg">
                        <p className="text-xs text-amber-700 dark:text-amber-300">
                            Comment banks are disabled because the comment has been manually edited.
                        </p>
                    </div>
                )}
                <div className="flex flex-col gap-4">
                    {renderGroupSection('Common', ccgBeforeSCG, commonSelections, handleCommonSelection, true)}
                    {overrideGroupsMapped.length > 0 && renderGroupSection('Subject Overrides', overrideGroupsMapped, selections, handleSelection, false)}
                    {renderGroupSection('Subject', pureScgGroupsMapped, selections, handleSelection, false)}
                    {renderGroupSection('Common', ccgAfterSCG, commonSelections, handleCommonSelection, true)}
                </div>
            </div>

            <div className="bg-white dark:bg-gray-900 rounded-xl p-6 border border-[#f0f2f4] dark:border-gray-800 shadow-sm">
                <h3 className="text-[#111418] dark:text-white text-sm font-bold uppercase tracking-wider mb-4">Selected Codes</h3>
                <div className="flex flex-wrap gap-2">
                     {/* Common group selections */}
                     {Object.entries(commonSelections).map(([gid, oid]) => {
                         const group = commonGroups?.find(g => g.id === gid);
                         const option = group?.CommonCommentOption.find(o => o.id === oid);
                         if (!group || !option) return null;
                         return (
                            <span key={gid} className="bg-green-100 dark:bg-green-900/30 text-green-700 dark:text-green-300 px-2 py-1 rounded text-xs font-medium border border-green-200 dark:border-green-800">
                                #{group.name}{option.code}
                            </span>
                         );
                     })}
                     {/* Subject group selections */}
                     {Object.entries(selections).map(([gid, oid]) => {
                         const group = groups.find(g => g.id === gid);
                         const option = group?.CommentOption.find(o => o.id === oid);
                         if (!group || !option) return null;
                         return (
                            <span key={gid} className="bg-blue-100 dark:bg-blue-900/30 text-blue-700 dark:text-blue-300 px-2 py-1 rounded text-xs font-medium border border-blue-200 dark:border-blue-800">
                                #{group.name}{option.code}
                            </span>
                         );
                     })}
                     {Object.keys(selections).length === 0 && Object.keys(commonSelections).length === 0 && <span className="text-gray-400 text-xs italic">No selection</span>}
                </div>
            </div>
        </aside>

        {/* Editor Section */}
        <div className="flex-1 flex flex-col bg-white dark:bg-gray-900 rounded-xl border border-[#f0f2f4] dark:border-gray-800 shadow-sm overflow-hidden min-h-[500px]">
            <div className="px-8 py-5 border-b border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center bg-gray-50/50 dark:bg-gray-800/50">
                <div className="flex items-center gap-3">
                    <h3 className="text-[#111418] dark:text-white text-lg font-bold leading-tight">Report Preview & Editor</h3>
                    <CommentStatusBadge status={checkStatus} size="sm" />
                    {commentBanksDisabled && (
                        <button
                            onClick={() => setShowRevertModal(true)}
                            className="flex items-center gap-1 px-2 py-1 text-xs font-medium text-amber-600 hover:text-amber-700 hover:bg-amber-50 dark:hover:bg-amber-900/20 rounded transition-colors"
                            title="Revert to generated comment"
                        >
                            <span className="material-symbols-outlined text-sm">undo</span>
                            Revert
                        </button>
                    )}
                </div>
                <div className="flex items-center gap-4">
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
                        disabled={checkStatus === 'required_check' || checkStatus === 'checked_rejected'}
                        className={`flex items-center gap-1 transition-colors ${
                            checkStatus === 'required_check' || checkStatus === 'checked_rejected'
                                ? 'text-gray-300 dark:text-gray-600 cursor-not-allowed'
                                : 'text-[#617289] dark:text-gray-400 hover:text-primary'
                        }`}
                        title={checkStatus === 'required_check' || checkStatus === 'checked_rejected' ? 'Comment must be approved before copying' : 'Copy to clipboard'}
                    >
                        <span className="material-symbols-outlined text-lg">{copied ? 'check' : 'content_copy'}</span>
                    </button>
                </div>
            </div>

            {/* Rejection Feedback Banner */}
            {checkStatus === 'checked_rejected' && assignment.checkNote && (
                <div className="mx-8 mt-4 p-4 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-lg">
                    <div className="flex items-start gap-3">
                        <span className="material-symbols-outlined text-red-500 mt-0.5">warning</span>
                        <div>
                            <p className="text-sm font-bold text-red-700 dark:text-red-400">Changes Requested</p>
                            <p className="text-sm text-red-600 dark:text-red-300 mt-1">{assignment.checkNote}</p>
                        </div>
                    </div>
                </div>
            )}

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

                </div>
            </div>

            {/* HoD Review Panel */}
            {isHoD && checkStatus === 'required_check' && (
                <div className="mx-8 mb-4 p-4 bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-800 rounded-lg">
                    <div className="flex flex-col gap-4">
                        <div className="flex items-center gap-2">
                            <span className="material-symbols-outlined text-amber-600">rate_review</span>
                            <p className="text-sm font-bold text-amber-700 dark:text-amber-400">Review Required</p>
                        </div>
                        <p className="text-sm text-amber-600 dark:text-amber-300">
                            This comment has been manually edited and requires your review before it can be finalized.
                        </p>
                        <div className="flex flex-col gap-3">
                            <textarea
                                className="w-full p-3 text-sm rounded-lg border border-amber-300 dark:border-amber-700 bg-white dark:bg-gray-900 text-gray-900 dark:text-gray-100 placeholder-gray-400 focus:ring-2 focus:ring-amber-400 focus:border-transparent resize-none"
                                placeholder="If rejecting, provide feedback for the teacher..."
                                rows={2}
                                value={rejectionNote}
                                onChange={(e) => setRejectionNote(e.target.value)}
                                disabled={isReviewing}
                            />
                            <div className="flex gap-3">
                                <button
                                    onClick={handleApprove}
                                    disabled={isReviewing}
                                    className="flex-1 flex items-center justify-center gap-2 px-4 py-2.5 bg-green-600 hover:bg-green-700 disabled:bg-green-400 text-white font-medium rounded-lg transition-colors"
                                >
                                    <span className="material-symbols-outlined text-lg">check_circle</span>
                                    {isReviewing ? 'Processing...' : 'Approve'}
                                </button>
                                <button
                                    onClick={handleReject}
                                    disabled={isReviewing || !rejectionNote.trim()}
                                    className="flex-1 flex items-center justify-center gap-2 px-4 py-2.5 bg-red-600 hover:bg-red-700 disabled:bg-red-400 text-white font-medium rounded-lg transition-colors"
                                >
                                    <span className="material-symbols-outlined text-lg">cancel</span>
                                    {isReviewing ? 'Processing...' : 'Reject'}
                                </button>
                            </div>
                            <p className="text-xs text-amber-500 dark:text-amber-400 italic">
                                Note: A reason is required when rejecting a comment.
                            </p>
                        </div>
                    </div>
                </div>
            )}

            <div className="px-8 py-4 bg-primary/5 dark:bg-primary/10 border-t border-[#f0f2f4] dark:border-gray-800 flex justify-between items-center">
                <div className="flex items-center gap-2 text-primary">
                    <span className="material-symbols-outlined text-base">lightbulb</span>
                    <span className="text-xs font-semibold">Tip: Changes are saved automatically.</span>
                </div>

            </div>
        </div>

        {/* Revert Confirmation Modal */}
        <ConfirmModal
            isOpen={showRevertModal}
            title="Revert to Generated Comment?"
            message="This will discard your manual edits and regenerate the comment from the selected codes. This action cannot be undone."
            confirmLabel={isReverting ? 'Reverting...' : 'Revert Comment'}
            cancelLabel="Keep Edits"
            confirmVariant="danger"
            onConfirm={handleRevert}
            onCancel={() => setShowRevertModal(false)}
        />
    </div>
  );
}
