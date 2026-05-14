export const runtime = 'nodejs';
export const maxDuration = 300;

import { NextRequest } from 'next/server';
import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';
import type { DbSubject, DbCommentGroup, DbCommentOption, DbCommonCommentGroup, DbCommonCommentOption } from '@/lib/types/db';
import type { AuditEvent, SpagAuditEntry, StandardsAuditEntry, AuditPdfData } from '@/lib/audit/types';
import { buildSampleReports } from '@/lib/audit/build-reports';
import { substituteVariables } from '@/lib/audit/substitute-variables';
import { callSpagWebhook, callStandardsWebhook } from '@/lib/audit/webhook-calls';
import { renderAuditPdf } from '@/lib/audit/generate-pdf';

export async function GET(
  _request: NextRequest,
  { params }: { params: Promise<{ subjectId: string }> }
) {
  const { subjectId } = await params;

  // ── Auth ──────────────────────────────────────────────────────────────────
  const session = await getServerSession(authOptions);
  if (!session?.user?.id) {
    return new Response('Unauthorized', { status: 401 });
  }

  // Verify the user is an admin or the HoD of this specific subject
  if (!session.user.roles?.includes('admin')) {
    const { rows: accessRows } = await pool.query(
      `SELECT 1 FROM "Subject" s
       JOIN "_SubjectToUser" su ON su."A" = s.id
       WHERE s.id = $1 AND su."B" = $2`,
      [subjectId, session.user.id]
    );
    if (accessRows.length === 0) {
      return new Response('Forbidden', { status: 403 });
    }
  }

  // ── Fetch subject ─────────────────────────────────────────────────────────
  const { rows: subjectRows } = await pool.query<DbSubject>(
    `SELECT * FROM "Subject" WHERE id = $1`,
    [subjectId]
  );
  if (subjectRows.length === 0) {
    return new Response('Not Found', { status: 404 });
  }
  const subject = subjectRows[0];

  // ── Fetch comment groups + options (non-linked only) ──────────────────────
  const { rows: groupRows } = await pool.query<DbCommentGroup>(
    `SELECT * FROM "CommentGroup" WHERE "subjectId" = $1 AND "isLinked" = false ORDER BY "displayOrder" ASC`,
    [subjectId]
  );

  type GroupWithOptions = {
    id: string; name: string; title: string; isLinked: boolean;
    options: { id: string; code: string; text: string; displayOrder: number }[];
  };

  const subjectGroups: GroupWithOptions[] = [];
  if (groupRows.length > 0) {
    const groupIds = groupRows.map(g => g.id);
    const { rows: optRows } = await pool.query<DbCommentOption>(
      `SELECT * FROM "CommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
      [groupIds]
    );
    for (const g of groupRows) {
      subjectGroups.push({
        id: g.id,
        name: g.name,
        title: g.title,
        isLinked: g.isLinked,
        options: optRows.filter(o => o.groupId === g.id),
      });
    }
  }

  // ── Fetch common comment groups + options ─────────────────────────────────
  // Fetch ALL CCGs (including linked ones) so that linked CCG slots in the
  // format template can be detected and overridden by subject-specific groups
  // with the same name (e.g. "Units Studied" per subject).
  const { rows: ccgRows } = await pool.query<DbCommonCommentGroup>(
    `SELECT * FROM "CommonCommentGroup" ORDER BY "displayOrder" ASC`
  );
  const commonGroups: GroupWithOptions[] = [];
  if (ccgRows.length > 0) {
    // Only fetch options for non-linked CCGs; linked CCGs have no options of their own
    const nonLinkedCcgIds = ccgRows.filter(g => !g.isLinked).map(g => g.id);
    const ccoRows = nonLinkedCcgIds.length > 0
      ? (await pool.query<DbCommonCommentOption>(
          `SELECT * FROM "CommonCommentOption" WHERE "groupId" = ANY($1::text[]) ORDER BY "displayOrder" ASC`,
          [nonLinkedCcgIds]
        )).rows
      : [];
    for (const g of ccgRows) {
      commonGroups.push({
        id: g.id,
        name: g.name,
        title: g.title,
        isLinked: g.isLinked,
        options: ccoRows.filter(o => o.groupId === g.id),
      });
    }
  }

  // ── Fetch ignored words for this teacher ──────────────────────────────────
  const { rows: ignoredRows } = await pool.query<{ word: string }>(
    `SELECT word FROM "IgnoredWord" WHERE "teacherId" = $1`,
    [session.user.id]
  );
  const ignoredWords = new Set(ignoredRows.map(r => r.word.toLowerCase()));

  // ── Collect all comment options for SPAG Phase 1 ─────────────────────────
  // Subject groups whose name matches a CCG name override that CCG slot.
  // We SPAG-check the subject group's options (the real content) and skip
  // the CCG's placeholder options to avoid duplicates in Section 1.
  const subjectGroupNames = new Set(subjectGroups.map(g => g.name));
  const overriddenCcgNames = new Set(
    commonGroups.filter(g => subjectGroupNames.has(g.name)).map(g => g.name)
  );

  const allComments: { code: string; text: string; groupName: string }[] = [];
  for (const g of subjectGroups) {
    for (const opt of g.options) {
      allComments.push({ code: opt.code, text: opt.text, groupName: g.title });
    }
  }
  for (const g of commonGroups) {
    if (overriddenCcgNames.has(g.name)) continue; // subject group handles this slot
    for (const opt of g.options) {
      allComments.push({ code: opt.code, text: opt.text, groupName: g.title });
    }
  }

  // ── Fetch format template + subject format ────────────────────────────────
  const { rows: settingRows } = await pool.query<{ value: string }>(
    `SELECT value FROM "AppSetting" WHERE key = 'comment_format_template'`
  );
  const formatTemplate = settingRows[0]?.value || '';
  const subjectFormat = subject.commentFormat || null;

  // ── Build sample reports ──────────────────────────────────────────────────
  const { reports, untestedItems } = buildSampleReports(
    subjectGroups, commonGroups, subject.title ?? subject.code, formatTemplate, subjectFormat
  );

  // ── Stream ────────────────────────────────────────────────────────────────
  const encoder = new TextEncoder();

  const stream = new ReadableStream({
    async start(controller) {
      // EventSource requires SSE format: "data: <json>\n\n"
      const send = (event: AuditEvent) => {
        controller.enqueue(encoder.encode(`data: ${JSON.stringify(event)}\n\n`));
      };

      try {
        send({ type: 'init', totalComments: allComments.length, totalReports: reports.length });

        // ── Phase 1: SPAG — check every individual comment option ──────────
        const spagEntries: SpagAuditEntry[] = [];

        for (const comment of allComments) {
          const substituted = substituteVariables(comment.text, subject.title ?? subject.code);
          const { passed, errors } = await callSpagWebhook(substituted, ignoredWords);
          const entry: SpagAuditEntry = {
            code: comment.code,
            groupName: comment.groupName,
            rawText: substituted,
            passed,
            errors,
          };
          spagEntries.push(entry);
          send({ type: 'spag', code: comment.code, groupName: comment.groupName, passed, errors });
        }

        send({ type: 'spag_done' });

        // ── Phase 2: Standards — check every sample report ─────────────────
        const standardsReports: StandardsAuditEntry[] = [];
        let passedReports = 0;

        for (const report of reports) {
          const { passed, failures, failureDetails } = await callStandardsWebhook(report.assembledText);
          if (passed) passedReports++;

          // Use groupTitle as key so the PDF can show human-readable "GroupTitle: Code" pairs
          const codes: Record<string, string> = {};
          for (const [, sel] of Object.entries(report.selections)) {
            codes[sel.groupTitle] = sel.code;
          }

          standardsReports.push({
            reportIndex: report.reportIndex,
            label: report.label,
            codes,
            passed,
            failures,
            assembledText: report.assembledText,
            failureDetails,
          });

          send({
            type: 'standards',
            reportIndex: report.reportIndex,
            codes,
            passed,
            failures,
          });
        }

        send({ type: 'standards_done' });
        send({ type: 'untested', items: untestedItems });

        // ── Generate PDF ───────────────────────────────────────────────────
        const pdfData: AuditPdfData = {
          subjectTitle: subject.title ?? subject.code,
          subjectCode: subject.code,
          generatedAt: new Date(),
          totalReports: reports.length,
          passedReports,
          spagEntries,
          standardsReports,
          untestedItems,
        };

        const pdfBuffer = await renderAuditPdf(pdfData);

        send({ type: 'complete', pdfBase64: pdfBuffer.toString('base64') });
      } catch (err) {
        send({ type: 'error', message: err instanceof Error ? err.message : String(err) });
      } finally {
        controller.close();
      }
    },
  });

  return new Response(stream, {
    headers: {
      'Content-Type': 'text/event-stream',
      'Cache-Control': 'no-cache, no-transform',
      'Connection': 'keep-alive',
      'X-Accel-Buffering': 'no',
    },
  });
}
