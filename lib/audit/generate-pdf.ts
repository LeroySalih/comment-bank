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
import type { StandardsRuleKey } from '@/lib/types/ai-check';

// ── All standards rules in display order ─────────────────────────────────────

const ALL_RULES: { key: StandardsRuleKey; label: string }[] = [
  { key: 'TargetWordCountMet',          label: 'Word Count' },
  { key: 'Formatting',                  label: 'Formatting' },
  { key: 'TerminologyCorrect',          label: 'Terminology' },
  { key: 'JargonFree',                  label: 'Jargon Free' },
  { key: 'ToneBalanced',                label: 'Tone Balanced' },
  { key: 'CourseOverviewIncluded',      label: 'Course Overview' },
  { key: 'SocialSkillsIncluded',        label: 'Social Skills' },
  { key: 'CollaborationIncluded',       label: 'Collaboration' },
  { key: 'BehaviourIncluded',           label: 'Behaviour' },
  { key: 'ParentalSupportIncluded',     label: 'Parental Support' },
];

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
  statValue:      { ...STAT_VALUE_BASE, color: 'white' },
  statValueGreen: { ...STAT_VALUE_BASE, color: '#86efac' },
  statValueRed:   { ...STAT_VALUE_BASE, color: '#fca5a5' },
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
    marginTop: 10,
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
    width: 36,
  },
  commentCodeFail: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    color: '#dc2626',
    width: 36,
  },
  commentText: {
    fontSize: 9,
    color: '#374151',
    flex: 1,
    lineHeight: 1.5,
  },
  spagError: {
    fontSize: 8,
    color: '#dc2626',
    paddingLeft: 48,
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
    paddingVertical: 14,
    paddingHorizontal: 32,
    marginBottom: 16,
  },
  reportPageHeaderPassed: {
    backgroundColor: '#14532d',
  },
  reportPageHeaderFailed: {
    backgroundColor: '#7f1d1d',
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
    fontSize: 9,
    color: '#111827',
    lineHeight: 1.6,
    borderLeftWidth: 3,
    borderLeftColor: '#d1d5db',
    paddingLeft: 12,
    marginBottom: 16,
  },
  rulesLabel: {
    fontSize: 8,
    fontFamily: 'Helvetica-Bold',
    color: '#374151',
    textTransform: 'uppercase',
    letterSpacing: 0.5,
    marginBottom: 6,
  },
  rulesGrid: {
    flexDirection: 'column',
    gap: 4,
  },
  ruleRow: {
    flexDirection: 'row',
    alignItems: 'flex-start',
    paddingVertical: 5,
    paddingHorizontal: 8,
    borderRadius: 4,
  },
  ruleRowPassed: {
    backgroundColor: '#f0fdf4',
  },
  ruleRowFailed: {
    backgroundColor: '#fef2f2',
  },
  ruleStatus: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    width: 18,
  },
  ruleStatusPassed: {
    color: '#16a34a',
  },
  ruleStatusFailed: {
    color: '#dc2626',
  },
  ruleName: {
    fontSize: 9,
    fontFamily: 'Helvetica-Bold',
    width: 130,
  },
  ruleNamePassed: {
    color: '#166534',
  },
  ruleNameFailed: {
    color: '#991b1b',
  },
  ruleFeedback: {
    fontSize: 8,
    color: '#6b7280',
    flex: 1,
    lineHeight: 1.4,
  },
  ruleFeedbackFailed: {
    color: '#b91c1c',
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

  // ── Page 1: Summary + Section 1 — All Comments ─────────────────────────────
  const mainPage = React.createElement(
    Page,
    { size: 'A4', style: styles.page },
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
          React.createElement(Text, { style: styles.statLabel }, 'Sample Reports')
        ),
        React.createElement(View, { style: styles.statBlock },
          React.createElement(Text, {
            style: data.passedReports === data.totalReports ? styles.statValueGreen : styles.statValueRed,
          }, `${passRate}%`),
          React.createElement(Text, { style: styles.statLabel }, 'Standards Passed')
        ),
        React.createElement(View, { style: styles.statBlock },
          React.createElement(Text, {
            style: spagFailures.length === 0 ? styles.statValueGreen : styles.statValueRed,
          }, String(spagFailures.length)),
          React.createElement(Text, { style: styles.statLabel }, 'SPAG Failures')
        ),
        React.createElement(View, { style: styles.statBlock },
          React.createElement(Text, {
            style: data.untestedItems.length > 0 ? styles.statValueRed : styles.statValueGreen,
          }, String(data.untestedItems.length)),
          React.createElement(Text, { style: styles.statLabel }, 'Untested Codes')
        )
      )
    ),
    React.createElement(
      View,
      { style: styles.body },
      React.createElement(Text, { style: styles.sectionLabel }, 'Section 1 — All Comments'),
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
                  `⚠ "${err.word}": ${err.message}`
                )
              )
            )
          )
        )
      ),
      data.untestedItems.length > 0
        ? React.createElement(Text, { style: styles.untestedNote },
            `⚠ ${data.untestedItems.length} code(s) not included in any sample report: ` +
            data.untestedItems.map(u => `${u.code} (${u.groupName})`).join(', ')
          )
        : null
    )
  );

  // ── One page per sample report — Section 2 ─────────────────────────────────
  const reportPages = data.standardsReports.map(report => {
    const headerStyle = {
      ...styles.reportPageHeader,
      ...(report.passed ? styles.reportPageHeaderPassed : styles.reportPageHeaderFailed),
    };

    return React.createElement(
      Page,
      { key: report.reportIndex, size: 'A4', style: styles.page },
      // Header
      React.createElement(
        View,
        { style: headerStyle },
        React.createElement(Text, { style: styles.reportPageHeaderLabel },
          `Section 2 — Sample Report ${report.reportIndex + 1} of ${data.totalReports} · ${report.label}`
        ),
        React.createElement(Text, { style: styles.reportPageHeaderTitle },
          report.passed ? `✓ ${report.label} — All Standards Passed` : `✗ ${report.label} — Standards Failures`
        ),
        React.createElement(Text, { style: styles.reportPageHeaderCodes },
          `Codes: ${Object.entries(report.codes).map(([title, code]) => `${title}: ${code}`).join('  |  ')}`
        )
      ),
      // Body
      React.createElement(
        View,
        { style: styles.reportBody },
        // Assembled text
        React.createElement(Text, { style: styles.reportTextLabel }, 'Assembled Comment'),
        React.createElement(Text, { style: styles.reportText }, report.assembledText),
        // Standards rules — all of them
        React.createElement(Text, { style: styles.rulesLabel }, 'Standards Audit Results'),
        React.createElement(
          View,
          { style: styles.rulesGrid },
          ...ALL_RULES.map(({ key, label }) => {
            const failed = report.failures.includes(key);
            const detail = report.failureDetails[key];
            const instances = detail?.instances ?? [];
            const feedbackText = failed
              ? instances.length > 0
                ? instances.join(' · ')
                : 'Did not meet this standard'
              : 'Passed';

            return React.createElement(
              View,
              { key, style: { ...styles.ruleRow, ...(failed ? styles.ruleRowFailed : styles.ruleRowPassed) } },
              React.createElement(Text, {
                style: { ...styles.ruleStatus, ...(failed ? styles.ruleStatusFailed : styles.ruleStatusPassed) },
              }, failed ? '✗' : '✓'),
              React.createElement(Text, {
                style: { ...styles.ruleName, ...(failed ? styles.ruleNameFailed : styles.ruleNamePassed) },
              }, label),
              React.createElement(Text, {
                style: { ...styles.ruleFeedback, ...(failed ? styles.ruleFeedbackFailed : {}) },
              }, feedbackText)
            );
          })
        )
      )
    );
  });

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
