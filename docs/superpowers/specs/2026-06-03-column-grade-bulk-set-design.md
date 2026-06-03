# Design: Column-Grade Bulk Set

**Date:** 2026-06-03  
**Branch:** worktree-column-grade

## Overview

Add a second sticky header row to the Class Matrix table containing "set all" buttons for each comment group column. Clicking a button sets that code for every pupil in the class that does not already have a code assigned in that group. Pupils with an existing code are unaffected.

Applies to all comment group types: subject-specific (SCG) and common (CCG) — both before and after SCG. Linked/read-only groups get no buttons.

---

## Architecture

The page remains a server component. A new thin `ClassMatrixClient` client component wraps the table and provides a React context for coordinating bulk updates between the button row and individual student rows.

### Files changed / created

| File | Change |
|------|--------|
| `app/actions.ts` | Add `bulkSetColumnCode` server action |
| `app/class/[classId]/page.tsx` | Wrap table in `ClassMatrixClient`; add second header row |
| `app/class/[classId]/_components/ClassMatrixClient.tsx` | New — context provider + bulk handler |
| `app/class/[classId]/_components/ClassMatrixContext.tsx` | New — context definition + types |
| `app/class/[classId]/_components/BulkCodeButtonRow.tsx` | New — second sticky header row |
| `app/class/[classId]/_components/StudentMatrixRow.tsx` | Register/unregister row handlers with context |

---

## Server Action

```ts
bulkSetColumnCode(
  classId: string,
  groupId: string,
  code: string,
  groupType: 'subject' | 'common'
): Promise<{ success: boolean; updatedAssignmentIds: string[]; error?: string }>
```

**Logic:**
1. Fetch all active assignment IDs for `classId`
2. Query `PupilCode` (SCG) or `CommonPupilCode` (CCG) to find assignments that already have a code for `groupId`
3. Compute the set difference: assignments with no existing code
4. Bulk insert with `INSERT ... ON CONFLICT DO NOTHING` for safety
5. Return `updatedAssignmentIds` (only the ones actually written)
6. Revalidate `/class/[classId]`

---

## Second Header Row (`BulkCodeButtonRow`)

A second `<tr>` inside `<thead>`, sticky below the first header row (`top-[header-height]`). Layout mirrors the first row:

- **Fixed columns** (Pupil Name, Gender, Status, EoY Level, Target, Actions): empty `<th>` cells with matching widths
- **Comment group columns**: one `<th>` per group containing a row of option buttons (e.g. High / Medium / Low)
- **Linked groups**: empty cell — no buttons
- Buttons are visually distinct from the data-row buttons (smaller, outlined style) to signal "apply to all"

---

## React Context (`ClassMatrixContext`)

```ts
type RowHandlers = {
  setCode: (groupId: string, code: string) => void          // SCG
  setCommonCode: (groupId: string, code: string) => void    // CCG
}

type ClassMatrixContextValue = {
  registerRow: (assignmentId: string, handlers: RowHandlers) => void
  unregisterRow: (assignmentId: string) => void
  applyBulkCode: (groupId: string, code: string, groupType: 'subject' | 'common') => void
}
```

---

## Client Wrapper (`ClassMatrixClient`)

- Holds a `Map<assignmentId, RowHandlers>` ref (not state — no re-render on register/unregister)
- `applyBulkCode(groupId, code, groupType)`:
  1. Calls `bulkSetColumnCode(classId, groupId, code, groupType)`
  2. On success, iterates `updatedAssignmentIds` and calls the matching row handler
  3. On error, surfaces a toast/alert

---

## `StudentMatrixRow` changes

- On mount: `registerRow(assignment.id, { setCode, setCommonCode })`
- On unmount: `unregisterRow(assignment.id)`
- `setCode` calls `setSelections` for the given groupId
- `setCommonCode` calls `setCommonSelections` for the given groupId
- Rows where `commentBanksDisabled` is true are still registered but their handlers are no-ops (consistent with existing per-cell disabled behaviour)

---

## Error Handling

- If `bulkSetColumnCode` returns `{ success: false }`, show an `alert()` (consistent with existing error handling in `StudentMatrixRow`)
- Partial success is not possible — the action either writes all missing rows or fails entirely (single transaction)

---

## Out of Scope

- Bulk-clearing codes (set to null)
- Undo/revert for bulk actions
- Visual indicator showing how many pupils were updated
