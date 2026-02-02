'use client';

import { useState, useEffect, useRef } from 'react';
import { useRouter } from 'next/navigation';
import { parseComment, countWords } from '@/lib/utils';
import { updateAssignmentCode, updateAssignmentCommentText, revertAssignmentComment } from '@/app/actions';
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
  CommentOption: CommentOption[];
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
  Pupil: Pupil;
  PupilCode: PupilCode[];
  finalComment?: string | null;
  eoyLevel?: string | null;
  targetLevel?: string | null;
  checkStatus?: string;
  checkNote?: string | null;
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
}

export default function CommentEditor({ assignment, subject, groups, isHoD = false }: CommentEditorProps) {
  const router = useRouter();
  const [selections, setSelections] = useState<Record<string, string>>({}); // groupId -> optionId
  const [preview, setPreview] = useState('');
  const [copied, setCopied] = useState(false);
  const [checkStatus, setCheckStatus] = useState(assignment.checkStatus || 'not_required');
  const [isManuallyEdited, setIsManuallyEdited] = useState(false); // Track if user manually edited
  const [showRevertModal, setShowRevertModal] = useState(false);
  const [isReverting, setIsReverting] = useState(false);

  // HoD review state
  const [rejectionNote, setRejectionNote] = useState('');
  const [isReviewing, setIsReviewing] = useState(false);

  // Initialize selections with assignment's pre-assigned codes
  useEffect(() => {
    const initialSelections: Record<string, string> = {};

    assignment.PupilCode.forEach(pc => {
        const group = groups.find(g => g.id === pc.groupId);
        if (group && pc.code) {
            const option = group.CommentOption.find(o => o.code === pc.code);
            if (option) {
                initialSelections[group.id] = option.id;
            }
        }
    });

    setSelections(initialSelections);
  }, [assignment, groups]);

  // Track if we've loaded the initial comment and should skip regeneration
  const hasLoadedInitialComment = useRef(false);
  const skipRegeneration = useRef(false); // Use ref instead of state to avoid race condition

  // Initialize preview from finalComment when it becomes available
  useEffect(() => {
    if (!hasLoadedInitialComment.current && assignment.finalComment) {
      setPreview(assignment.finalComment);
      hasLoadedInitialComment.current = true;
      skipRegeneration.current = true; // Don't regenerate after loading saved comment
    }
  }, [assignment.finalComment]); // Run when finalComment changes

  // Generate Preview from selections (only when selections change and not manually edited)
  useEffect(() => {
    // Skip if user has manually typed in the textarea
    if (isManuallyEdited) {
      return;
    }

    // Skip if we loaded from a saved comment (check ref immediately)
    if (skipRegeneration.current) {
      return;
    }

    // Skip on first render if we have a saved comment (let the mount effect handle it)
    if (!hasLoadedInitialComment.current && assignment.finalComment) {
      return;
    }

    const parts: string[] = [];

    // 1. Studied Comment (Subject Intro)
    if (subject.studiedComment) {
      parts.push(subject.studiedComment);
    }

    // 2. Get text for all groups in order
    const sortedGroups = [...groups].sort((a,b) => ((a as any).displayOrder || 0) - ((b as any).displayOrder || 0));

    const getOptText = (group: CommentGroup) => {
        const selectedOptionId = selections[group.id];
        if (!selectedOptionId) return "";
        return group.CommentOption.find(o => o.id === selectedOptionId)?.text || "";
    };

    const groupTexts: string[] = [];
    for (const group of sortedGroups) {
        const text = getOptText(group);
        if (text) {
            groupTexts.push(text);
        }
    }

    // Join all group texts with spaces
    if (groupTexts.length > 0) {
        parts.push(groupTexts.join(" "));
    }

    // Combine all parts with paragraph breaks
    const combined = parts.join("\n\n");

    setPreview(parseComment(
      combined,
      assignment.Pupil.firstName,
      assignment.Pupil.gender,
      subject.title || '',
      assignment.Class?.year,
      assignment.eoyLevel,
      assignment.targetLevel
    ));

  }, [selections, subject, groups, assignment, isManuallyEdited]);

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

  const copyToClipboard = () => {
    navigator.clipboard.writeText(preview);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleTextChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setPreview(e.target.value);
    setIsManuallyEdited(true); // Mark as manually edited to prevent regeneration
  };

  const handleTextBlur = async () => {
    // Only save and update status if user manually edited the text
    // Don't trigger for programmatic changes from code selection
    if (!isManuallyEdited) {
      return;
    }

    // Update local state immediately for responsive UI
    const previousStatus = checkStatus;
    setCheckStatus('required_check');

    try {
      const result = await updateAssignmentCommentText(assignment.id, preview);

      if (!result.success) {
        console.error('Failed to save comment:', result.error);
        // Revert local state if save failed
        setCheckStatus(previousStatus);
        alert('Failed to save comment: ' + (result.error || 'Unknown error'));
      }
      // Don't call router.refresh() - it causes the preview to regenerate from codes
      // The local state already has the correct values
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
        // Reset local state
        setIsManuallyEdited(false);
        skipRegeneration.current = false;
        hasLoadedInitialComment.current = false;
        setCheckStatus('not_required');
        setShowRevertModal(false);
        // Refresh to get fresh data and regenerate comment
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
  const targetWordCount = 100; // Configurable or hardcoded for now
  const percent = Math.min(100, Math.round((wordCount / targetWordCount) * 100));
  const dashArray = 100;
  const dashOffset = 100 - (percent / 100 * 100); // SVG stroke-dashoffset calculation

  // Disable comment banks if comment has been manually edited
  const commentBanksDisabled = isManuallyEdited || skipRegeneration.current;

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
                <div className="flex flex-col gap-2">
                    {groups.map(group => {
                        const selectedOptionId = selections[group.id];
                        const selectedOption = group.CommentOption.find(o => o.id === selectedOptionId);
                        const isSelected = !!selectedOption;

                        return (
                            <div key={group.id} className={`flex flex-col gap-2 px-3 py-3 rounded-lg transition-colors border border-transparent ${commentBanksDisabled ? 'cursor-not-allowed' : 'cursor-pointer'} ${isSelected ? 'bg-primary/5 border-primary/10' : commentBanksDisabled ? '' : 'hover:bg-[#f0f2f4] dark:hover:bg-gray-800'}`}>
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
                                    {group.CommentOption.map(opt => (
                                        <button
                                            key={opt.id}
                                            onClick={() => !commentBanksDisabled && handleSelection(group.id, opt.id)}
                                            disabled={commentBanksDisabled}
                                            className={`text-xs px-2 py-1 rounded border ${selections[group.id] === opt.id ? 'bg-primary text-white border-primary' : 'bg-white text-gray-600 border-gray-200'} ${commentBanksDisabled ? 'cursor-not-allowed opacity-50' : 'hover:border-gray-300'}`}
                                            title={commentBanksDisabled ? 'Comment banks are locked' : opt.text}
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
                         const option = group?.CommentOption.find(o => o.id === oid);
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
