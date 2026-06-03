'use client';

import { useRef, useCallback } from 'react';
import { ClassMatrixContext, RowHandlers } from './ClassMatrixContext';
import { bulkSetColumnCode } from '@/app/actions';

interface ClassMatrixClientProps {
  classId: string;
  children: React.ReactNode;
}

export default function ClassMatrixClient({ classId, children }: ClassMatrixClientProps) {
  const rowHandlers = useRef<Map<string, RowHandlers>>(new Map());

  const registerRow = useCallback((assignmentId: string, handlers: RowHandlers) => {
    rowHandlers.current.set(assignmentId, handlers);
  }, []);

  const unregisterRow = useCallback((assignmentId: string) => {
    rowHandlers.current.delete(assignmentId);
  }, []);

  const applyBulkCode = useCallback(async (
    groupId: string,
    code: string,
    groupType: 'subject' | 'common'
  ) => {
    const result = await bulkSetColumnCode(classId, groupId, code, groupType);
    if (!result.success) {
      alert('Failed to apply bulk code: ' + (result.error || 'Unknown error'));
      return;
    }
    for (const assignmentId of result.updatedAssignmentIds) {
      const handlers = rowHandlers.current.get(assignmentId);
      if (!handlers) continue;
      if (groupType === 'subject') {
        handlers.setCode(groupId, code);
      } else {
        handlers.setCommonCode(groupId, code);
      }
    }
  }, [classId]);

  return (
    <ClassMatrixContext.Provider value={{ registerRow, unregisterRow, applyBulkCode }}>
      {children}
    </ClassMatrixContext.Provider>
  );
}
