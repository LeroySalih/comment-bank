// lib/audit/generate-pdf.ts

import React from 'react';
import {
  Document,
  Page,
  Text,
  View,
  StyleSheet,
  renderToBuffer,
} from '@react-pdf/renderer';
import type { AuditPdfData } from './types';

// ── Styles ────────────────────────────────────────────────────────────────────

const styles = StyleSheet.create({
  page: {
    fontFamily: 'Helvetica',
    fontSize: 10,
    color: '#1a1a1a',
    paddingTop: 0,
    paddingBottom: 32,
    paddingHorizontal: 0,
  },
  // Dark header banner
  header: {
    backgroundColor: '#1e3a5f',
    color: 'white',
    paddingVertical: 20,
    paddingHorizontal: 32,
    marginBottom: 24,
  },
  headerTitle: {
    fontSize: 18,
    fontFamily: 'Helvetica-Bold',
    color: 'white',
    marginBottom: 4,
  },
  headerSubtitle: {
    fontSize: 9,
    color: '#9ab3c8',
    marginBottom: 12,
  },
  headerStats: {
    flexDirection: 'row',
    gap: 24,
  },
  statBlock: {
    flexDirection: 'column',
  },
  statValue: {
    fontSize: 20,
    fontFamily: 'Helvetica-Bold',
    color: 'white',
  },
  statValueGreen: {
    fontSize: 20,
    fontFamily: 'Helvetica-Bold',
    color: '#86efac',
  },
  statValueRed: {
    fontSize: 20,
    fontFamily: 'Helvetica-Bold',
    color: '#fca5a5',
  },
  statLabel: {
    fontSize: 8,
    color: '#9ab3c8',
    marginTop: 2,
  },
  body: {
    paddingHorizontal: 32,
  },
  sectionLabel: {
    fontSize: 8,
    fontFamily: 'Helvetica-Bold',
    color: '#6b7280',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: 8,
    marginTop: 16,
  },
  groupTitle: {
    fontSize: 10,
    fontFamily: 'Helvetica-Bold',
    color: '#1e3a5f',
    borderLeftWidth: 3,
    borderLeftColor: '#1e3a5f',
    paddingLeft: 8,
    marginBottom: 4,
    marginTop: 8,
  },
  commentRow: {
    flexDirection: 'row',
    marginBottom: 4,
    paddingLeft: 12,
  },
  commentCode: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#2563eb',
    width: 28,
  },
  commentCodeFail: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#dc2626',
    width: 28,
  },
  commentText: {
    fontSize: 9,
    color: '#374151',
    flex: 1,
  },
  spagError: {
    fontSize: 8,
    color: '#dc2626',
    paddingLeft: 40,
    marginBottom: 2,
  },
  failureCard: {
    borderWidth: 1,
    borderColor: '#fecaca',
    borderRadius: 4,
    marginBottom: 6,
    overflow: 'hidden',
  },
  failureCardHeader: {
    backgroundColor: '#fee2e2',
    paddingVertical: 5,
    paddingHorizontal: 8,
    flexDirection: 'row',
    gap: 8,
  },
  failureCardHeaderText: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#991b1b',
  },
  failureCardBody: {
    paddingVertical: 5,
    paddingHorizontal: 8,
  },
  failureRule: {
    fontSize: 9,
    color: '#dc2626',
    marginBottom: 2,
  },
  untestedNote: {
    fontSize: 9,
    color: '#92400e',
    backgroundColor: '#fef3c7',
    borderRadius: 4,
    padding: 8,
    marginTop: 8,
  },
});

// ── Document component ────────────────────────────────────────────────────────

function buildAuditDocument(data: AuditPdfData) {
  const passRate = data.totalReports > 0
    ? Math.round((data.passedReports / data.totalReports) * 100)
    : 0;

  const spagFailures = data.spagEntries.filter(e => !e.passed);

  // Group spag entries by group name
  const groupedEntries = new Map<string, typeof data.spagEntries>();
  for (const entry of data.spagEntries) {
    const arr = groupedEntries.get(entry.groupName) ?? [];
    arr.push(entry);
    groupedEntries.set(entry.groupName, arr);
  }

  const dateStr = data.generatedAt.toLocaleDateString('en-GB', {
    day: 'numeric', month: 'long', year: 'numeric',
  });

  return React.createElement(
    Document,
    null,
    React.createElement(
      Page,
      { size: 'A4', style: styles.page },
      // Header
      React.createElement(
        View,
        { style: styles.header },
        React.createElement(Text, { style: styles.headerTitle },
          `${data.subjectCode} — Comment Bank Audit`
        ),
        React.createElement(Text, { style: styles.headerSubtitle },
          `${data.subjectTitle} · Generated ${dateStr}`
        ),
        React.createElement(
          View,
          { style: styles.headerStats },
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, { style: styles.statValue }, String(data.totalReports)),
            React.createElement(Text, { style: styles.statLabel }, 'Reports')
          ),
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, { style: styles.statValueGreen }, `${passRate}%`),
            React.createElement(Text, { style: styles.statLabel }, 'Passed')
          ),
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, { style: styles.statValueRed }, String(spagFailures.length)),
            React.createElement(Text, { style: styles.statLabel }, 'SPAG Failures')
          ),
          React.createElement(View, { style: styles.statBlock },
            React.createElement(Text, {
              style: data.untestedItems.length > 0 ? styles.statValueRed : styles.statValueGreen,
            }, String(data.untestedItems.length)),
            React.createElement(Text, { style: styles.statLabel }, 'Untested')
          )
        )
      ),
      // Body
      React.createElement(
        View,
        { style: styles.body },
        // Section 1 — Comments Audited
        React.createElement(Text, { style: styles.sectionLabel }, 'Section 1 — Comments Audited'),
        ...[...groupedEntries.entries()].map(([groupName, entries]) =>
          React.createElement(
            View,
            { key: groupName },
            React.createElement(Text, { style: styles.groupTitle }, groupName),
            ...entries.map(entry =>
              React.createElement(
                View,
                { key: entry.code },
                React.createElement(
                  View,
                  { style: styles.commentRow },
                  React.createElement(Text, {
                    style: entry.passed ? styles.commentCode : styles.commentCodeFail,
                  }, entry.passed ? entry.code : `${entry.code} ✗`),
                  React.createElement(Text, { style: styles.commentText }, entry.rawText)
                ),
                ...entry.errors.map((err, i) =>
                  React.createElement(Text, { key: i, style: styles.spagError },
                    `  ⚠ "${err.word}": ${err.message}`
                  )
                )
              )
            )
          )
        ),

        // Section 2 — Failed Standards Reports
        data.standardsFailures.length > 0
          ? React.createElement(
            View,
            null,
            React.createElement(Text, { style: styles.sectionLabel }, 'Section 2 — Failed Standards Reports'),
            ...data.standardsFailures.map(failure =>
              React.createElement(
                View,
                { key: failure.reportIndex, style: styles.failureCard },
                React.createElement(
                  View,
                  { style: styles.failureCardHeader },
                  React.createElement(Text, { style: styles.failureCardHeaderText },
                    `Report #${failure.reportIndex + 1}`
                  ),
                  React.createElement(Text, { style: styles.failureCardHeaderText },
                    Object.values(failure.codes).join(', ')
                  )
                ),
                React.createElement(
                  View,
                  { style: styles.failureCardBody },
                  ...failure.failures.map(rule =>
                    React.createElement(Text, { key: rule, style: styles.failureRule },
                      `✗ ${rule}`
                    )
                  )
                )
              )
            )
          )
          : React.createElement(Text, { style: { ...styles.sectionLabel, color: '#16a34a' } },
            '✓ All standards reports passed'
          ),

        // Untested warning
        data.untestedItems.length > 0
          ? React.createElement(Text, { style: styles.untestedNote },
            `⚠ ${data.untestedItems.length} comment code(s) were not included in any of the 50 sample reports: ` +
            data.untestedItems.map(u => `${u.code} (${u.groupName})`).join(', ')
          )
          : null
      )
    )
  );
}

// ── Public render function ────────────────────────────────────────────────────

export async function renderAuditPdf(data: AuditPdfData): Promise<Buffer> {
  return renderToBuffer(buildAuditDocument(data));
}
