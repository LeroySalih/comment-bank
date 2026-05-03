'use client';

import { useEffect, useRef, useState, useCallback } from 'react';
import { createPortal } from 'react-dom';
import type { AuditEvent } from '@/lib/audit/types';

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
      untestedCount: number;
      pdfUrl: string;
    }
  | { name: 'error'; message: string };

interface AuditModalProps {
  subjectId: string;
  subjectTitle: string;
  isOpen: boolean;
  onClose: () => void;
}

export default function AuditModal({
  subjectId,
  subjectTitle,
  isOpen,
  onClose,
}: AuditModalProps) {
  const [phase, setPhase] = useState<Phase>({ name: 'idle' });
  const esRef = useRef<EventSource | null>(null);
  const spagFailRef = useRef(0);
  const totalCommentsRef = useRef(0);
  const totalReportsRef = useRef(0);
  const passedReportsRef = useRef(0);
  const untestedCountRef = useRef(0);

  const closeStream = useCallback(() => {
    esRef.current?.close();
    esRef.current = null;
  }, []);

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
        if (!event.passed) spagFailRef.current++;
        setPhase(prev =>
          prev.name === 'phase1'
            ? { ...prev, checkedComments: prev.checkedComments + 1, currentLabel: `${event.groupName}: ${event.code}` }
            : prev
        );
        break;

      case 'spag_done':
        setPhase(prev =>
          prev.name === 'phase1'
            ? { name: 'phase2', totalComments: prev.totalComments, totalReports: totalReportsRef.current, checkedReports: 0, currentLabel: 'Building sample reports…' }
            : prev
        );
        break;

      case 'standards':
        if (event.passed) passedReportsRef.current++;
        setPhase(prev =>
          prev.name === 'phase2'
            ? { ...prev, checkedReports: prev.checkedReports + 1, currentLabel: `Report #${event.reportIndex + 1}` }
            : prev
        );
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
          untestedCount: untestedCountRef.current,
          pdfUrl: event.pdfUrl,
        });
        break;

      case 'error':
        closeStream();
        setPhase({ name: 'error', message: event.message });
        break;
    }
  }

  useEffect(() => {
    if (!isOpen) {
      closeStream();
      setPhase({ name: 'idle' });
      spagFailRef.current = 0;
      totalCommentsRef.current = 0;
      totalReportsRef.current = 0;
      passedReportsRef.current = 0;
      untestedCountRef.current = 0;
      return;
    }

    const es = new EventSource(`/api/subjects/${subjectId}/audit`);
    esRef.current = es;

    es.onmessage = (e: MessageEvent) => {
      const event = JSON.parse(e.data) as AuditEvent;
      handleEvent(event);
    };

    es.onerror = () => {
      closeStream();
      setPhase({ name: 'error', message: 'Connection to audit service lost.' });
    };

    return () => {
      closeStream();
    };
  }, [isOpen, subjectId, closeStream]);

  // Escape key
  useEffect(() => {
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isOpen) {
        closeStream();
        onClose();
      }
    };
    document.addEventListener('keydown', handleKey);
    return () => document.removeEventListener('keydown', handleKey);
  }, [isOpen, onClose, closeStream]);

  if (!isOpen || typeof document === 'undefined') return null;

  const handleDownload = () => {
    if (phase.name !== 'complete') return;
    const a = document.createElement('a');
    a.href = phase.pdfUrl;
    a.download = '';
    a.click();
  };

  const handleCancel = () => {
    closeStream();
    onClose();
  };

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-center justify-center">
      <div className="absolute inset-0 bg-black/50 backdrop-blur-sm" onClick={handleCancel} />
      <div className="relative bg-white dark:bg-gray-900 rounded-xl shadow-2xl max-w-md w-full mx-4 p-6">
        <h3 className="text-lg font-bold text-gray-900 dark:text-white mb-1">
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
                <span>Phase 1: SPAG checking comments</span>
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
                    width: phase.name === 'phase1' && phase.totalComments > 0
                      ? `${(phase.checkedComments / phase.totalComments) * 100}%`
                      : '100%',
                  }}
                />
              </div>
            </div>

            {/* Phase 2 bar */}
            <div className={phase.name === 'phase1' ? 'opacity-40' : ''}>
              <div className="flex justify-between text-xs text-gray-600 dark:text-gray-400 mb-1">
                <span>Phase 2: Standards checking reports</span>
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
                    width: phase.name === 'phase2' && phase.totalReports > 0
                      ? `${(phase.checkedReports / phase.totalReports) * 100}%`
                      : '0%',
                  }}
                />
              </div>
            </div>

            <p className="text-xs text-gray-400 dark:text-gray-500 truncate">
              {phase.name === 'phase1' ? phase.currentLabel : phase.currentLabel}
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
                onClick={onClose}
                className="flex-1 px-4 py-2 text-sm bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                Retry
              </button>
            </div>
          </div>
        )}
      </div>
    </div>,
    document.body
  );
}
