'use client';

import { createContext, useContext } from 'react';

export type RowHandlers = {
  setCode: (groupId: string, code: string) => void;
  setCommonCode: (groupId: string, code: string) => void;
};

export type ClassMatrixContextValue = {
  registerRow: (assignmentId: string, handlers: RowHandlers) => void;
  unregisterRow: (assignmentId: string) => void;
  applyBulkCode: (groupId: string, code: string, groupType: 'subject' | 'common') => void;
};

export const ClassMatrixContext = createContext<ClassMatrixContextValue>({
  registerRow: () => {},
  unregisterRow: () => {},
  applyBulkCode: () => {},
});

export function useClassMatrix() {
  return useContext(ClassMatrixContext);
}
