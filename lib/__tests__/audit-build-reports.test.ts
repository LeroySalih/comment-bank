import { describe, it, expect } from 'vitest';
import { buildSampleReports } from '../audit/build-reports';

type Option = { id: string; code: string; text: string; displayOrder: number };
type Group = { id: string; name: string; title: string; isLinked: boolean; options: Option[] };

function makeGroup(id: string, name: string, title: string, codes: string[], isLinked = false): Group {
  return {
    id,
    name,
    title,
    isLinked,
    options: codes.map((code, i) => ({
      id: `${id}-opt-${i}`,
      code,
      text: `${title} ${code} comment text`,
      displayOrder: i + 1,
    })),
  };
}

const COMMON: Group[] = [
  makeGroup('cg1', 'Effort', 'Effort', ['Excellent', 'Good', 'Poor']),
];

describe('buildSampleReports', () => {
  it('generates exactly 3 base reports for 3-option groups', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'Skills', 'Skills', ['High', 'Med', 'Low']),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing');
    expect(reports.length).toBe(3);
  });

  it('labels reports correctly', () => {
    const groups = [makeGroup('g1', 'Knowledge', 'Knowledge', ['High', 'Med', 'Low'])];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing');
    expect(reports[0].label).toBe('All High');
    expect(reports[1].label).toBe('All Medium');
    expect(reports[2].label).toBe('All Low');
  });

  it('skips isLinked groups', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'LinkedGroup', 'LinkedGroup', ['A', 'B'], true),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing');
    // Only g1 should appear in selections, not g2
    for (const r of reports) {
      expect(r.selections['g2']).toBeUndefined();
    }
  });

  it('includes common group options in assembled text', () => {
    const groups = [makeGroup('g1', 'Knowledge', 'Knowledge', ['High'])];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing');
    expect(reports.length).toBeGreaterThan(0);
    // Common group text should be in the assembled output
    expect(reports[0].assembledText).toContain('Effort');
  });

  it('generates "all high" report as first report', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'Skills', 'Skills', ['High', 'Med', 'Low']),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing');
    // First report: lowest displayOrder (index 0) from each group = "High"
    expect(reports[0].selections['g1'].code).toBe('High');
    expect(reports[0].selections['g2'].code).toBe('High');
  });

  it('generates coverage fill reports for 2-option groups', () => {
    // 2 options per group: All High covers opt[0], All Medium covers opt[0] again (middle=0),
    // All Low covers opt[1]. Coverage fill should add one more report to cover opt[1] for g2.
    const groups = [
      makeGroup('g1', 'Knowledge', 'Knowledge', ['A', 'B']),
    ];
    const { reports, untestedItems } = buildSampleReports(groups, COMMON, 'Computing');
    expect(untestedItems.length).toBe(0);
  });

  it('reports zero untested when all codes are covered', () => {
    // 1 group × 1 option = fully covered in 1 report
    const groups = [makeGroup('g1', 'Knowledge', 'Knowledge', ['OnlyOption'])];
    const { untestedItems } = buildSampleReports(groups, COMMON, 'Computing');
    expect(untestedItems.length).toBe(0);
  });

  it('common groups vary by tier', () => {
    const groups = [makeGroup('g1', 'Knowledge', 'Knowledge', ['High', 'Med', 'Low'])];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing');
    // All High: common group cg1 should use first option (Excellent)
    expect(reports[0].selections['cg1'].code).toBe('Excellent');
    // All Low: common group cg1 should use last option (Poor)
    expect(reports[2].selections['cg1'].code).toBe('Poor');
  });
});
