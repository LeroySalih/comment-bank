'use client';

import { useEffect, useRef, useState, useCallback } from 'react';
import { createPortal } from 'react-dom';
import type { AuditEvent } from '@/lib/audit/types';

// ── Phase state ───────────────────────────────────────────────────────────────

type Phase =
  | { name: 'idle' }
  | {
      name: 'phase1';
      totalComments: number;
      checkedComments: number;
      currentLabel: string;
    }
  | {
      name: 'phase2';
      totalComments: number;
      totalReports: number;
      checkedReports: number;
      currentLabel: string;
    }
  | {
      name: 'complete';
      totalReports: number;
      passedReports: number;
      spagFailures: number;
      spagFailedComments: number;
      untestedCount: number;
      pdfBase64: string;
    }
  | { name: 'error'; message: string };

// ── Props ─────────────────────────────────────────────────────────────────────

interface AuditModalProps {
  subjectId: string;
  subjectTitle: string;
  isOpen: boolean;
  onClose: () => void;
}

// ── Component ─────────────────────────────────────────────────────────────────

export default function AuditModal({
  subjectId,
  subjectTitle,
  isOpen,
  onClose,
}: AuditModalProps) {
  const [mounted, setMounted] = useState(false);
  const [phase, setPhase] = useState<Phase>({ name: 'idle' });
  // retryKey increments on Retry to force the stream useEffect to re-run
  const [retryKey, setRetryKey] = useState(0);

  const dialogRef = useRef<HTMLDialogElement | null>(null);
  const esRef = useRef<EventSource | null>(null);
  const prevFocusRef = useRef<Element | null>(null);

  // Accumulated refs — safe to read in the handleEvent closure because refs
  // are always current (no stale-closure risk).
  const totalCommentsRef = useRef(0);
  const totalReportsRef = useRef(0);
  const passedReportsRef = useRef(0);
  const spagFailRef = useRef(0);          // total SPAG error events
  const spagFailedCodesRef = useRef<Set<string>>(new Set()); // distinct failing comment codes
  const untestedCountRef = useRef(0);

  // ── Stream management ─────────────────────────────────────────────────────

  const closeStream = useCallback(() => {
    esRef.current?.close();
    esRef.current = null;
  }, []);

  const resetAccumulators = useCallback(() => {
    totalCommentsRef.current = 0;
    totalReportsRef.current = 0;
    passedReportsRef.current = 0;
    spagFailRef.current = 0;
    spagFailedCodesRef.current = new Set();
    untestedCountRef.current = 0;
  }, []);

  // ── Suppress portal until client has mounted (prevents hydration mismatch) ──
  useEffect(() => { setMounted(true); }, []);

  // ── Open stream when isOpen becomes true (or on retry) ───────────────────

  useEffect(() => {
    if (!isOpen) {
      closeStream();
      setPhase({ name: 'idle' });
      resetAccumulators();
      return;
    }

    // Reset for a fresh run (handles retries)
    setPhase({ name: 'idle' });
    resetAccumulators();

    const es = new EventSource(`/api/subjects/${subjectId}/audit`);
    esRef.current = es;

    // handleEvent defined inside the effect so it closes over the freshly-created
    // `es` instance — no stale closures, all mutable state goes through refs.
    function handleEvent(event: AuditEvent) {
      switch (event.type) {
        case 'init':
          totalCommentsRef.current = event.totalComments;
          totalReportsRef.current = event.totalReports;
          setPhase({
            name: 'phase1',
            totalComments: event.totalComments,
            checkedComments: 0,
            currentLabel: 'Starting…',
          });
          break;

        case 'spag':
          if (!event.passed) {
            spagFailRef.current++;
            spagFailedCodesRef.current.add(event.code);
          }
          setPhase(prev =>
            prev.name === 'phase1'
              ? {
                  ...prev,
                  checkedComments: prev.checkedComments + 1,
                  currentLabel: `${event.groupName}: ${event.code}`,
                }
              : prev
          );
          break;

        case 'spag_done':
          setPhase(prev =>
            prev.name === 'phase1'
              ? {
                  name: 'phase2',
                  totalComments: prev.totalComments,
                  totalReports: totalReportsRef.current,
                  checkedReports: 0,
                  currentLabel: 'Building sample reports…',
                }
              : prev
          );
          break;

        case 'standards':
          if (event.passed) passedReportsRef.current++;
          setPhase(prev =>
            prev.name === 'phase2'
              ? {
                  ...prev,
                  checkedReports: prev.checkedReports + 1,
                  currentLabel: `Report #${event.reportIndex + 1}`,
                }
              : prev
          );
          break;

        case 'standards_done':
          // No UI change needed — complete event follows immediately
          break;

        case 'untested':
          untestedCountRef.current = event.items.length;
          break;

        case 'complete':
          closeStream();
          setPhase({
            name: 'complete',
            totalReports: totalReportsRef.current,
            passedReports: passedReportsRef.current,
            spagFailures: spagFailRef.current,
            spagFailedComments: spagFailedCodesRef.current.size,
            untestedCount: untestedCountRef.current,
            pdfBase64: event.pdfBase64,
          });
          break;

        case 'error':
          closeStream();
          setPhase({ name: 'error', message: event.message });
          break;

        default:
          // Exhaustiveness guard — new event types will cause a TS error here
          break;
      }
    }

    es.onmessage = (e: MessageEvent) => {
      try {
        const event = JSON.parse(e.data) as AuditEvent;
        handleEvent(event);
      } catch {
        closeStream();
        setPhase({ name: 'error', message: 'Received malformed data from audit service.' });
      }
    };

    es.onerror = () => {
      // Only set error if the stream is still open (i.e. not already complete).
      // The browser fires onerror when the server closes the SSE connection normally
      // after a complete event — we must not overwrite the complete phase.
      if (esRef.current === null) return;
      closeStream();
      setPhase({ name: 'error', message: 'Connection to audit service lost.' });
    };

    return () => {
      closeStream();
    };
  }, [isOpen, subjectId, retryKey, closeStream, resetAccumulators]);

  // ── Focus management ──────────────────────────────────────────────────────

  useEffect(() => {
    if (isOpen) {
      prevFocusRef.current = document.activeElement;
      // showModal() promotes to the top layer and auto-focuses first focusable child
      dialogRef.current?.showModal();
    } else {
      dialogRef.current?.close();
      // Restore focus to the element that opened the modal
      if (prevFocusRef.current instanceof HTMLElement) {
        prevFocusRef.current.focus();
      }
    }
  }, [isOpen]);

  // ── Escape key (close the stream and dismiss) ─────────────────────────────

  useEffect(() => {
    if (!isOpen) return;

    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        closeStream();
        onClose();
      }
    };
    document.addEventListener('keydown', handleKey);
    return () => document.removeEventListener('keydown', handleKey);
  }, [isOpen, onClose, closeStream]);

  // ── Action handlers ───────────────────────────────────────────────────────

  const handleCancel = () => {
    closeStream();
    onClose();
  };

  const handleDownload = () => {
    if (phase.name !== 'complete') return;
    try {
      // Decode the base64 PDF that arrived in the SSE complete event — no second
      // HTTP request needed, so this works regardless of which server process
      // handled the original stream.
      const binary = atob(phase.pdfBase64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const blob = new Blob([bytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `audit-${subjectTitle.replace(/\s+/g, '-')}.pdf`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error('PDF download failed:', err);
    }
  };

  const handleRetry = () => {
    // Increment retryKey to re-trigger the stream useEffect while isOpen stays true
    setRetryKey(k => k + 1);
  };

  // ── Render ────────────────────────────────────────────────────────────────

  if (!mounted) return null;

  const headingId = `audit-modal-heading-${subjectId}`;

  return createPortal(
    <dialog
      ref={dialogRef}
      aria-modal="true"
      aria-labelledby={headingId}
      className="max-w-md w-full mx-4 p-0 rounded-xl shadow-2xl bg-white dark:bg-gray-900 backdrop:bg-black/50 backdrop:backdrop-blur-sm"
      onClose={onClose}
    >
      <div className="p-6">
        <h3 id={headingId} className="text-lg font-bold text-gray-900 dark:text-white mb-1">
          Comment Bank Audit
        </h3>
        <p className="text-sm text-gray-500 dark:text-gray-400 mb-6">{subjectTitle}</p>

        {phase.name === 'idle' && (
          <p className="text-sm text-gray-500">Starting audit…</p>
        )}

        {(phase.name === 'phase1' || phase.name === 'phase2') && (
          <div className="space-y-5">
            {/* Phase 1 bar */}
            <div>
              <div className="flex justify-between text-xs text-gray-600 dark:text-gray-400 mb-1">
                <span>Phase 1: SPAG checking individual comments</span>
                <span>
                  {phase.name === 'phase1' ? phase.checkedComments : phase.totalComments}
                  {' / '}
                  {phase.totalComments}
                </span>
              </div>
              <div className="bg-gray-200 dark:bg-gray-700 rounded-full h-2">
                <div
                  className="bg-blue-500 h-2 rounded-full transition-all duration-300"
                  style={{
                    width:
                      phase.name === 'phase1' && phase.totalComments > 0
                        ? `${(phase.checkedComments / phase.totalComments) * 100}%`
                        : '100%',
                  }}
                />
              </div>
            </div>

            {/* Phase 2 bar — greyed out until spag_done */}
            <div className={phase.name === 'phase1' ? 'opacity-40' : ''}>
              <div className="flex justify-between text-xs text-gray-600 dark:text-gray-400 mb-1">
                <span>Phase 2: Standards checking sample reports</span>
                <span>
                  {phase.name === 'phase2' ? phase.checkedReports : 0}
                  {' / '}
                  {phase.name === 'phase2' ? phase.totalReports : totalReportsRef.current}
                </span>
              </div>
              <div className="bg-gray-200 dark:bg-gray-700 rounded-full h-2">
                <div
                  className="bg-blue-500 h-2 rounded-full transition-all duration-300"
                  style={{
                    width:
                      phase.name === 'phase2' && phase.totalReports > 0
                        ? `${(phase.checkedReports / phase.totalReports) * 100}%`
                        : '0%',
                  }}
                />
              </div>
            </div>

            {/* Currently checking */}
            <p className="text-xs text-gray-400 dark:text-gray-500 truncate">
              {phase.currentLabel}
            </p>

            <div className="flex justify-end">
              <button
                onClick={handleCancel}
                className="px-4 py-2 text-sm text-gray-600 dark:text-gray-400 hover:text-gray-800 dark:hover:text-white transition-colors"
              >
                Cancel
              </button>
            </div>
          </div>
        )}

        {phase.name === 'complete' && (
          <div className="space-y-4">
            <div className="flex items-center gap-2 text-green-600 dark:text-green-400">
              <span className="material-symbols-outlined">check_circle</span>
              <span className="font-semibold">Audit Complete</span>
            </div>
            <ul className="text-sm text-gray-600 dark:text-gray-400 space-y-1">
              <li>{phase.totalReports} reports generated</li>
              <li>
                {phase.totalReports > 0
                  ? Math.round((phase.passedReports / phase.totalReports) * 100)
                  : 0}
                % passed standards checks
              </li>
              <li className={phase.spagFailures > 0 ? 'text-red-500' : ''}>
                {phase.spagFailures} SPAG {phase.spagFailures === 1 ? 'failure' : 'failures'}
                {phase.spagFailedComments > 0 && ` across ${phase.spagFailedComments} ${phase.spagFailedComments === 1 ? 'comment' : 'comments'}`}
              </li>
              <li className={phase.untestedCount > 0 ? 'text-amber-500' : ''}>
                {phase.untestedCount} comment {phase.untestedCount === 1 ? 'code' : 'codes'} untested
              </li>
            </ul>
            <div className="flex gap-3 pt-2">
              <button
                onClick={onClose}
                className="flex-1 px-4 py-2 text-sm border border-gray-300 dark:border-gray-600 rounded-lg text-gray-700 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-800 transition-colors"
              >
                Close
              </button>
              <button
                onClick={handleDownload}
                className="flex-1 px-4 py-2 text-sm bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors flex items-center justify-center gap-2"
              >
                <span className="material-symbols-outlined text-base">download</span>
                Download PDF
              </button>
            </div>
          </div>
        )}

        {phase.name === 'error' && (
          <div className="space-y-4">
            <div className="flex items-center gap-2 text-red-500">
              <span className="material-symbols-outlined">error</span>
              <span className="font-semibold">Audit Failed</span>
            </div>
            <p className="text-sm text-gray-600 dark:text-gray-400">{phase.message}</p>
            <div className="flex gap-3">
              <button
                onClick={onClose}
                className="flex-1 px-4 py-2 text-sm border border-gray-300 dark:border-gray-600 rounded-lg text-gray-700 dark:text-gray-300 hover:bg-gray-50 dark:hover:bg-gray-800 transition-colors"
              >
                Close
              </button>
              <button
                onClick={handleRetry}
                className="flex-1 px-4 py-2 text-sm bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                Retry
              </button>
            </div>
          </div>
        )}
      </div>
    </dialog>,
    document.body
  );
}
