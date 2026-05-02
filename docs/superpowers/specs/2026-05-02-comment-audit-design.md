# Comment Bank Audit — Design Spec

**Date:** 2026-05-02  
**Status:** Approved

---

## Overview

A subject-level audit tool that checks every individual comment for spelling/grammar errors and runs 50 structured sample reports through the standards checker. Results are delivered as a downloadable PDF. The audit is triggered from the subject admin page and runs inside a modal with live progress.

---

## Goals

- Catch SPAG errors in any comment option before they appear in pupil reports
- Validate that realistic combinations of comments produce standards-passing reports
- Surface which specific comment codes contribute to failures
- Flag any comment codes that weren't covered by the 50-report sample

---

## Data Pipeline

### Phase 1 — SPAG per comment

Fetch all `CommentOption` rows for the subject (all non-linked `CommentGroup` rows) plus all `CommonCommentOption` rows. Skip any group where `isLinked = true` (not yet implemented).

Before sending to the SPAG webhook, substitute the following fixed test values into template variables:

| Variable | Test value |
|---|---|
| `<Name>` | Alex |
| `<he/she>` | they |
| `<his/her>` | their |
| `<him/her>` | them |
| `<Subject>` | (subject title from DB) |
| `<Year>` | Year 10 |
| `<EoYLevel>` | 6 |
| `<TargetLevel>` | 7 |

Each substituted comment is sent individually to `SPAG_WEBHOOK_URL`. Results are accumulated as a map of `code → { passed, errors: SpagMatch[] }`.

The ignored words list (`IgnoredWord` table, scoped to the triggering teacher) is applied server-side as per the existing `requestSpagCheck` implementation.

### Phase 2 — Standards on 50 sample reports

Build up to 50 full report texts by selecting one `CommentOption` per non-linked subject group and appending a fixed representative selection from each `CommonCommentGroup` (first option by `displayOrder`).

Reports are assembled in this priority order until 50 are filled:

| Priority | Strategy | Description |
|---|---|---|
| 1 | **All High** | Highest `displayOrder` option from every group |
| 2 | **All Medium** | Middle option per group (floor of count/2) |
| 3 | **All Low** | Lowest `displayOrder` option per group |
| 4 | **Mostly High** | All High, except one group rotated to Medium — one report per subject group |
| 5 | **Mostly Low** | All Low, except one group rotated to High — one report per subject group |
| 6 | **Coverage fill** | Random combinations weighted to include option codes not yet seen in any report |

Each assembled text is sent to `STANDARDS_WEBHOOK_URL`. Results accumulate as an array of `{ reportIndex, codes: Record<groupId, code>, passed, failures: StandardsRuleKey[] }`.

After 50 reports, any `CommentOption.code` that did not appear in any report is collected as the **untested** list.

---

## API

### SSE Audit Route

```
GET /api/subjects/[subjectId]/audit
```

- Streams newline-delimited JSON events via `ReadableStream`
- Sets `export const maxDuration = 300` to handle large subjects
- Requires the user to be authenticated and to have admin access to the subject

**Event types:**

```ts
{ type: 'init';        totalComments: number; totalReports: number }
{ type: 'spag';        code: string; groupName: string; passed: boolean; errors: SpagMatch[] }
{ type: 'spag_done' }
{ type: 'standards';   reportIndex: number; codes: Record<string, string>; passed: boolean; failures: string[] }
{ type: 'standards_done' }
{ type: 'untested';    items: { code: string; groupName: string }[] }
{ type: 'complete';    pdfUrl: string }
{ type: 'error';       message: string }
```

### PDF Serve Route

```
GET /api/subjects/[subjectId]/audit/pdf?token=<token>
```

- Token is a random UUID generated at the end of the SSE handler
- Both routes must use `export const runtime = 'nodejs'` — the in-memory PDF store does not survive edge cold starts
- Stored in memory (server-side Map) with a 5-minute TTL
- Returns the PDF buffer with `Content-Type: application/pdf` and `Content-Disposition: attachment`
- Deleted from the in-memory store after first download

---

## PDF Generation

Uses `@react-pdf/renderer` (server-side, no headless browser). Generated at the end of the SSE handler before the `complete` event is emitted.

### Structure (Option B — dark header banner)

**Page 1**

- **Header banner** (dark navy): subject name, date generated, four key stats inline — Reports Generated / % Standards Passed / SPAG Failures / Untested Comments
- **Section 1 — Comments Audited**: grouped by `CommentGroup.title`, each option listed as `[code] — full comment text`. SPAG failures shown in red beneath the comment with the specific `SpagMatch.message`.
- **Section 2 — Failed Standards Reports**: one entry per failing report. Shows the comment codes used (one per group) and the list of failed `StandardsRuleKey` names.

**Page 2+ (auto)**

Overflow from Section 2 if there are many failures — `@react-pdf/renderer` handles pagination automatically.

---

## Modal UI

### Component

`components/AuditModal.tsx` — a native `<dialog>` element, consistent with existing modal patterns.

### Trigger

An **"Audit Comments"** button added to the subject admin page. Opens the modal and immediately begins the SSE stream.

### States

```
idle → running:phase1 → running:phase2 → complete → error
```

### Running view

```
Comment Bank Audit — [Subject Name]

Phase 1: SPAG checking individual comments
████████████████░░░░░░  24 / 31

Phase 2: Standards checking sample reports
░░░░░░░░░░░░░░░░░░░░░░   0 / 50   (greyed until phase 1 done)

Currently checking: "Alex has demonstrated…"

                        [Cancel]
```

- Phase 2 bar is rendered but visually greyed out until `spag_done` is received
- The "currently checking" line updates with each `spag` or `standards` event
- Cancel closes the `EventSource` and dismisses the modal (audit continues server-side until completion; PDF is simply not downloaded)

### Complete view

```
Audit Complete ✓

50 reports generated
84% passed standards checks
5 SPAG failures across 3 comments
3 comment codes untested

              [Download PDF]    [Close]
```

Download triggers a fetch to the short-lived PDF token URL and initiates browser download. Modal closes on Escape or the Close button.

### Error view

If an `error` event is received, the modal shows the error message with a Retry button that restarts the SSE stream from scratch.

---

## File Locations

| File | Purpose |
|---|---|
| `app/api/subjects/[subjectId]/audit/route.ts` | SSE streaming audit handler |
| `app/api/subjects/[subjectId]/audit/pdf/route.ts` | PDF token-gated download handler |
| `lib/audit/build-reports.ts` | Assembles the 50 sample report texts |
| `lib/audit/substitute-variables.ts` | Replaces template variables with fixed test values |
| `lib/audit/generate-pdf.ts` | `@react-pdf/renderer` PDF document component + render function |
| `lib/audit/pdf-store.ts` | In-memory token → PDF buffer store with TTL cleanup |
| `components/AuditModal.tsx` | Modal UI with SSE client and progress bars |
| `components/AuditButton.tsx` | Trigger button added to subject admin page |

---

## Dependencies

- `@react-pdf/renderer` — add to `package.json` (server-side PDF generation)

No other new dependencies. SPAG and Standards checks reuse the existing `requestSpagCheck` and `requestStandardsCheck` implementations from `lib/server-actions/ai-check.ts`, extracted into plain async functions that can be called outside a server action context.

---

## Out of Scope

- Linked comment groups (`isLinked = true`) — skipped entirely; noted in the PDF header
- Per-pupil audit (this is subject-level only)
- Storing audit results in the database (PDF is ephemeral)
- Scheduling recurring audits
