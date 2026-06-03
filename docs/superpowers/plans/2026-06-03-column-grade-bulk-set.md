# Column-Grade Bulk Set Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a second sticky header row to the Class Matrix table with buttons that bulk-set a comment code for all pupils in the class who don't already have one for that group.

**Architecture:** A new `bulkSetColumnCode` server action handles the DB write. A React context (`ClassMatrixContext`) connects the bulk-set button row to individual student rows so only un-set rows update their local state without a page reload.

**Tech Stack:** Next.js 14 App Router, React context, server actions, `pg` Pool, TypeScript, Tailwind CSS

---

## File Map

| File | Action |
|------|--------|
| `app/actions.ts` | Add `bulkSetColumnCode` |
| `app/class/[classId]/_components/ClassMatrixContext.tsx` | Create — context types + `createContext` |
| `app/class/[classId]/_components/ClassMatrixClient.tsx` | Create — context provider + `applyBulkCode` logic |
| `app/class/[classId]/_components/BulkCodeButtonRow.tsx` | Create — second sticky `<tr>` |
| `app/class/[classId]/_components/StudentMatrixRow.tsx` | Modify — register/unregister row handlers |
| `app/class/[classId]/page.tsx` | Modify — wrap table in `ClassMatrixClient`, add `BulkCodeButtonRow` |

---

## Task 1: Add `bulkSetColumnCode` server action

**Files:**
- Modify: `app/actions.ts`

- [ ] **Step 1: Add the action at the bottom of `app/actions.ts`**

```ts
export async function bulkSetColumnCode(
  classId: string,
  groupId: string,
  code: string,
  groupType: 'subject' | 'common'
): Promise<{ success: boolean; updatedAssignmentIds: string[]; error?: string }> {
  try {
    // Fetch all active assignment IDs for this class
    const { rows: assignmentRows } = await pool.query<{ id: string }>(
      `SELECT a.id FROM "Assignment" a
       JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
       WHERE a."classId" = $1 AND p."isActive" = true`,
      [classId]
    )
    const allIds = assignmentRows.map(r => r.id)
    if (allIds.length === 0) return { success: true, updatedAssignmentIds: [] }

    // Find which assignments already have a code for this group
    let alreadySetIds: string[] = []
    if (groupType === 'subject') {
      const { rows } = await pool.query<{ assignmentId: string }>(
        `SELECT "assignmentId" FROM "PupilCode"
         WHERE "assignmentId" = ANY($1::text[]) AND "groupId" = $2`,
        [allIds, groupId]
      )
      alreadySetIds = rows.map(r => r.assignmentId)
    } else {
      const { rows } = await pool.query<{ assignmentId: string }>(
        `SELECT "assignmentId" FROM "CommonPupilCode"
         WHERE "assignmentId" = ANY($1::text[]) AND "commonGroupId" = $2`,
        [allIds, groupId]
      )
      alreadySetIds = rows.map(r => r.assignmentId)
    }

    const alreadySetSet = new Set(alreadySetIds)
    const toUpdate = allIds.filter(id => !alreadySetSet.has(id))
    if (toUpdate.length === 0) return { success: true, updatedAssignmentIds: [] }

    // Bulk insert with ON CONFLICT DO NOTHING for safety
    if (groupType === 'subject') {
      for (const assignmentId of toUpdate) {
        await pool.query(
          `INSERT INTO "PupilCode" (id, "assignmentId", "groupId", code)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT ("assignmentId", "groupId") DO NOTHING`,
          [createId(), assignmentId, groupId, code]
        )
      }
    } else {
      for (const assignmentId of toUpdate) {
        await pool.query(
          `INSERT INTO "CommonPupilCode" (id, "assignmentId", "commonGroupId", code)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT ("assignmentId", "commonGroupId") DO NOTHING`,
          [createId(), assignmentId, groupId, code]
        )
      }
    }

    revalidatePath(`/class/${classId}`)

    return { success: true, updatedAssignmentIds: toUpdate }
  } catch (error) {
    console.error('Failed to bulk set column code:', error)
    return { success: false, updatedAssignmentIds: [], error: 'Database error' }
  }
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
cd /Users/leroysalih/nodejs/comment-bank/.claude/worktrees/column-grade
npx tsc --noEmit 2>&1 | head -20
```

Expected: no errors (or only pre-existing unrelated errors).

- [ ] **Step 3: Commit**

```bash
git add app/actions.ts
git commit -m "feat: add bulkSetColumnCode server action"
```

---

## Task 2: Create `ClassMatrixContext`

**Files:**
- Create: `app/class/[classId]/_components/ClassMatrixContext.tsx`

- [ ] **Step 1: Create the context file**

```tsx
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
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
npx tsc --noEmit 2>&1 | head -20
```

Expected: no new errors.

- [ ] **Step 3: Commit**

```bash
git add "app/class/[classId]/_components/ClassMatrixContext.tsx"
git commit -m "feat: add ClassMatrixContext"
```

---

## Task 3: Create `ClassMatrixClient`

**Files:**
- Create: `app/class/[classId]/_components/ClassMatrixClient.tsx`

- [ ] **Step 1: Create the client wrapper**

```tsx
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
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
npx tsc --noEmit 2>&1 | head -20
```

Expected: no new errors.

- [ ] **Step 3: Commit**

```bash
git add "app/class/[classId]/_components/ClassMatrixClient.tsx"
git commit -m "feat: add ClassMatrixClient context provider"
```

---

## Task 4: Create `BulkCodeButtonRow`

**Files:**
- Create: `app/class/[classId]/_components/BulkCodeButtonRow.tsx`

This component renders a second `<tr>` inside `<thead>` with bulk-set buttons for each comment group column.

- [ ] **Step 1: Create the component**

```tsx
'use client';

import { useState } from 'react';
import { useClassMatrix } from './ClassMatrixContext';

type Option = {
  id: string;
  code: string;
  text: string;
};

type Group = {
  id: string;
  name: string;
  isLinked?: boolean;
  CommentOption: Option[];
};

type CommonCommentGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  CommonCommentOption: Option[];
};

interface BulkCodeButtonRowProps {
  groups: Group[];
  commonGroupsBefore: CommonCommentGroup[];
  commonGroupsAfter: CommonCommentGroup[];
}

function BulkButton({
  groupId,
  code,
  groupType,
}: {
  groupId: string;
  code: string;
  groupType: 'subject' | 'common';
}) {
  const { applyBulkCode } = useClassMatrix();
  const [loading, setLoading] = useState(false);

  const handleClick = async () => {
    setLoading(true);
    await applyBulkCode(groupId, code, groupType);
    setLoading(false);
  };

  return (
    <button
      onClick={handleClick}
      disabled={loading}
      title={`Set all unset pupils to ${code}`}
      className={`px-2.5 py-1 text-xs font-bold rounded border transition-colors
        border-[#dbe0e6] dark:border-[#3a4454]
        text-[#617289] dark:text-gray-400
        hover:bg-gray-100 dark:hover:bg-[#2d3748]
        disabled:opacity-50 disabled:cursor-not-allowed`}
    >
      {loading ? '…' : code}
    </button>
  );
}

export default function BulkCodeButtonRow({
  groups,
  commonGroupsBefore,
  commonGroupsAfter,
}: BulkCodeButtonRowProps) {
  const emptyCellClass =
    'sticky top-[57px] z-40 px-6 py-2 bg-gray-50 dark:bg-[#151d28] border-b border-[#e5e7eb] dark:border-[#2d3748]';

  return (
    <tr>
      {/* Fixed columns — empty */}
      <th className={`${emptyCellClass} left-0 w-[240px] min-w-[240px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[240px] w-[80px] min-w-[80px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[320px] w-[140px] min-w-[140px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[460px] w-[100px] min-w-[100px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />
      <th className={`${emptyCellClass} left-[560px] w-[100px] min-w-[100px] shadow-[1px_0_0_0_rgba(229,231,235,1)] dark:shadow-[1px_0_0_0_rgba(45,55,72,1)]`} />

      {/* CCG before SCG */}
      {commonGroupsBefore.map((g) => (
        <th key={g.id} className={`${emptyCellClass} min-w-[200px]`}>
          {!g.isLinked && (
            <div className="flex gap-1">
              {g.CommonCommentOption.map((opt) => (
                <BulkButton key={opt.id} groupId={g.id} code={opt.code} groupType="common" />
              ))}
            </div>
          )}
        </th>
      ))}

      {/* Subject-specific groups */}
      {groups.map((g) => (
        <th key={g.id} className={`${emptyCellClass} min-w-[200px]`}>
          {!g.isLinked && (
            <div className="flex gap-1">
              {g.CommentOption.map((opt) => (
                <BulkButton key={opt.id} groupId={g.id} code={opt.code} groupType="subject" />
              ))}
            </div>
          )}
        </th>
      ))}

      {/* CCG after SCG */}
      {commonGroupsAfter.map((g) => (
        <th key={g.id} className={`${emptyCellClass} min-w-[200px]`}>
          {!g.isLinked && (
            <div className="flex gap-1">
              {g.CommonCommentOption.map((opt) => (
                <BulkButton key={opt.id} groupId={g.id} code={opt.code} groupType="common" />
              ))}
            </div>
          )}
        </th>
      ))}

      {/* Actions column — empty */}
      <th className={`${emptyCellClass} right-0 min-w-[120px]`} />
    </tr>
  );
}
```

- [ ] **Step 2: Verify TypeScript compiles**

```bash
npx tsc --noEmit 2>&1 | head -20
```

Expected: no new errors.

- [ ] **Step 3: Commit**

```bash
git add "app/class/[classId]/_components/BulkCodeButtonRow.tsx"
git commit -m "feat: add BulkCodeButtonRow component"
```

---

## Task 5: Update `StudentMatrixRow` to register with context

**Files:**
- Modify: `app/class/[classId]/_components/StudentMatrixRow.tsx`

- [ ] **Step 1: Add imports at the top of `StudentMatrixRow.tsx`**

After the existing imports, add:

```tsx
import { useEffect } from 'react';
import { useClassMatrix } from './ClassMatrixContext';
```

Note: `useState` is already imported — only add `useEffect`.

- [ ] **Step 2: Inside the component body, after the existing `useState` hooks, add registration**

After the line `const [isReverting, setIsReverting] = useState(false);`, add:

```tsx
  const { registerRow, unregisterRow } = useClassMatrix();

  useEffect(() => {
    registerRow(assignment.id, {
      setCode: (groupId, code) => {
        if (commentBanksDisabled) return;
        setSelections(prev => ({ ...prev, [groupId]: code }));
      },
      setCommonCode: (groupId, code) => {
        if (commentBanksDisabled) return;
        setCommonSelections(prev => ({ ...prev, [groupId]: code }));
      },
    });
    return () => unregisterRow(assignment.id);
  }, [assignment.id, commentBanksDisabled, registerRow, unregisterRow]);
```

- [ ] **Step 3: Verify TypeScript compiles**

```bash
npx tsc --noEmit 2>&1 | head -20
```

Expected: no new errors.

- [ ] **Step 4: Commit**

```bash
git add "app/class/[classId]/_components/StudentMatrixRow.tsx"
git commit -m "feat: register StudentMatrixRow handlers with ClassMatrixContext"
```

---

## Task 6: Update `page.tsx` to wire everything together

**Files:**
- Modify: `app/class/[classId]/page.tsx`

- [ ] **Step 1: Add imports at the top of `page.tsx`**

After the existing import of `StudentMatrixRow`, add:

```tsx
import ClassMatrixClient from './_components/ClassMatrixClient';
import BulkCodeButtonRow from './_components/BulkCodeButtonRow';
```

- [ ] **Step 2: Wrap the table container in `ClassMatrixClient`**

Find the line:
```tsx
        <div className="flex-1 overflow-auto p-6">
```

Replace it with:
```tsx
        <ClassMatrixClient classId={classId}>
        <div className="flex-1 overflow-auto p-6">
```

And find the closing `</div>` that matches this wrapper (the one after `</div>` for the table container), and add `</ClassMatrixClient>` after it. The end of the `return` block should look like:

```tsx
        </div>
        </ClassMatrixClient>
    </main>
```

- [ ] **Step 3: Add `BulkCodeButtonRow` as a second row inside `<thead>`**

Inside `<thead>`, after the closing `</tr>` of the first (existing) header row, add:

```tsx
                            <BulkCodeButtonRow
                                groups={groups}
                                commonGroupsBefore={commonGroupsBefore}
                                commonGroupsAfter={commonGroupsAfter}
                            />
```

- [ ] **Step 4: Adjust the first header row's sticky top to `top-0` and bump `BulkCodeButtonRow` cells to `top-[57px]`**

The first header row `<th>` elements already use `sticky top-0`. The `BulkCodeButtonRow` cells use `top-[57px]` (hardcoded in Task 4). No changes needed here — verify the existing header row `<th>` cells already have `sticky top-0` in their className (they do — check the file).

- [ ] **Step 5: Verify TypeScript compiles**

```bash
npx tsc --noEmit 2>&1 | head -20
```

Expected: no new errors.

- [ ] **Step 6: Commit**

```bash
git add "app/class/[classId]/page.tsx"
git commit -m "feat: integrate ClassMatrixClient and BulkCodeButtonRow into class page"
```

---

## Task 7: Manual verification

- [ ] **Step 1: Start the dev server**

```bash
npm run dev
```

- [ ] **Step 2: Navigate to a class with multiple pupils**

Open `http://localhost:3000` and navigate to a class matrix page (e.g. `/class/<classId>`).

- [ ] **Step 3: Verify the second header row appears**

You should see a second sticky row below the column headers containing small outlined buttons (e.g. High / Medium / Low) for each comment group column. Fixed columns (Pupil Name, Gender, etc.) should be empty in this row.

- [ ] **Step 4: Clear one pupil's code for a group, then click the bulk button for that code**

Select a pupil that already has a code set and note it. Clear one other pupil's code for the same group. Click the matching bulk button. Expected: only the cleared pupil gets that code set. The pupil who already had a code is unchanged.

- [ ] **Step 5: Verify linked groups show no buttons**

If any group column is linked (shows a badge instead of buttons), confirm its cell in the bulk row is empty.

- [ ] **Step 6: Commit a note if no changes needed**

If no bugs found, no commit needed. If fixes are required, fix and commit with:

```bash
git commit -m "fix: <description of fix>"
```
