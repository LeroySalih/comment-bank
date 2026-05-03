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

const STAT_VALUE_BASE = { fontSize: 20, fontFamily: 'Helvetica-Bold' } as const;

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
  statValue: { ...STAT_VALUE_BASE, color: 'white' },
  statValueGreen: { ...STAT_VALUE_BASE, color: '#86efac' },
  statValueRed: { ...STAT_VALUE_BASE, color: '#fca5a5' },
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
  untestedNote: {
    fontSize: 9,
    color: '#92400e',
    backgroundColor: '#fef3c7',
    borderRadius: 4,
    padding: 8,
    marginTop: 8,
  },
  // ── Per-report page styles ──────────────────────────────────────────────────
  reportPageHeader: {
    backgroundColor: '#1e3a5f',
    paddingVertical: 14,
    paddingHorizontal: 32,
    marginBottom: 20,
  },
  reportPageHeaderLabel: {
    fontSize: 8,
    color: '#9ab3c8',
    marginBottom: 2,
  },
  reportPageHeaderTitle: {
    fontSize: 14,
    fontFamily: 'Helvetica-Bold',
    color: 'white',
    marginBottom: 4,
  },
  reportPageHeaderCodes: {
    fontSize: 8,
    color: '#c7d9ea',
  },
  reportBody: {
    paddingHorizontal: 32,
  },
  reportTextLabel: {
    fontSize: 8,
    fontFamily: 'Helvetica-Bold',
    color: '#6b7280',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: 6,
  },
  reportText: {
    fontSize: 10,
    color: '#111827',
    lineHeight: 1.6,
    borderLeftWidth: 3,
    borderLeftColor: '#d1d5db',
    paddingLeft: 12,
    marginBottom: 20,
  },
  failedRulesLabel: {
    fontSize: 8,
    fontFamily: 'Helvetica-Bold',
    color: '#991b1b',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: 8,
  },
  ruleItem: {
    marginBottom: 10,
    borderLeftWidth: 2,
    borderLeftColor: '#fca5a5',
    paddingLeft: 10,
  },
  ruleName: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#dc2626',
    marginBottom: 3,
  },
  ruleInstance: {
    fontSize: 8,
    color: '#6b7280',
    marginBottom: 1,
    paddingLeft: 8,
  },
  ruleNoInstances: {
    fontSize: 8,
    color: '#9ca3af',
    fontStyle: 'italic',
    paddingLeft: 8,
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

  // ── Page 1: Summary + Section 1 ────────────────────────────────────────────
  const mainPage = React.createElement(
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

      // Note about Section 2
      data.standardsFailures.length > 0
        ? React.createElement(Text, {
            style: { ...styles.sectionLabel, marginTop: 24, color: '#dc2626' },
          },
            `Section 2 — ${data.standardsFailures.length} Failed Standards Report${data.standardsFailures.length > 1 ? 's' : ''} (see following pages)`
          )
        : React.createElement(Text, { style: { ...styles.sectionLabel, color: '#16a34a', marginTop: 24 } },
            '✓ All standards reports passed'
          ),

      // Untested warning
      data.untestedItems.length > 0
        ? React.createElement(Text, { style: styles.untestedNote },
            `⚠ ${data.untestedItems.length} comment code(s) were not included in any of the ${data.totalReports} sample reports: ` +
            data.untestedItems.map(u => `${u.code} (${u.groupName})`).join(', ')
          )
        : null
    )
  );

  // ── One page per failed standards report ────────────────────────────────────
  const reportPages = data.standardsFailures.map(failure =>
    React.createElement(
      Page,
      { key: failure.reportIndex, size: 'A4', style: styles.page },
      // Mini header
      React.createElement(
        View,
        { style: styles.reportPageHeader },
        React.createElement(Text, { style: styles.reportPageHeaderLabel },
          `Section 2 — Failed Standards Report`
        ),
        React.createElement(Text, { style: styles.reportPageHeaderTitle },
          `Report #${failure.reportIndex + 1}`
        ),
        React.createElement(Text, { style: styles.reportPageHeaderCodes },
          `Comment codes: ${Object.values(failure.codes).join(', ')}`
        )
      ),
      // Report body
      React.createElement(
        View,
        { style: styles.reportBody },
        // Full assembled text
        React.createElement(Text, { style: styles.reportTextLabel }, 'Assembled Report'),
        React.createElement(Text, { style: styles.reportText }, failure.assembledText),
        // Failed rules
        React.createElement(Text, { style: styles.failedRulesLabel },
          `Failed Rules (${failure.failures.length})`
        ),
        ...failure.failures.map(rule => {
          const detail = failure.failureDetails[rule];
          const instances = detail?.instances ?? [];
          return React.createElement(
            View,
            { key: rule, style: styles.ruleItem },
            React.createElement(Text, { style: styles.ruleName }, `✗ ${rule}`),
            instances.length > 0
              ? instances.map((inst, i) =>
                  React.createElement(Text, { key: i, style: styles.ruleInstance },
                    `• ${inst}`
                  )
                )
              : React.createElement(Text, { style: styles.ruleNoInstances },
                  'No additional detail available'
                )
          );
        })
      )
    )
  );

  return React.createElement(
    Document,
    null,
    mainPage,
    ...reportPages
  );
}

// ── Public render function ────────────────────────────────────────────────────

export async function renderAuditPdf(data: AuditPdfData): Promise<Buffer> {
  return renderToBuffer(buildAuditDocument(data));
}
