import { describe, it, expect } from 'vitest';
import { buildSampleReports } from '../audit/build-reports';

type Option = { id: string; code: string; text: string; displayOrder: number };
type Group = { id: string; title: string; isLinked: boolean; options: Option[] };

function makeGroup(id: string, title: string, codes: string[], isLinked = false): Group {
  return {
    id,
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
  makeGroup('cg1', 'Effort', ['Excellent', 'Good', 'Poor']),
];

describe('buildSampleReports', () => {
  it('returns at most maxReports reports', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'Skills', ['High', 'Med', 'Low']),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 5);
    expect(reports.length).toBeLessThanOrEqual(5);
  });

  it('skips isLinked groups', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'LinkedGroup', ['A', 'B'], true),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 50);
    // Only g1 should appear in selections, not g2
    for (const r of reports) {
      expect(r.selections['g2']).toBeUndefined();
    }
  });

  it('includes common group options in assembled text', () => {
    const groups = [makeGroup('g1', 'Knowledge', ['High'])];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 50);
    expect(reports.length).toBeGreaterThan(0);
    // Common group text should be in the assembled output
    expect(reports[0].assembledText).toContain('Effort');
  });

  it('generates "all high" report as first report', () => {
    const groups = [
      makeGroup('g1', 'Knowledge', ['High', 'Med', 'Low']),
      makeGroup('g2', 'Skills', ['High', 'Med', 'Low']),
    ];
    const { reports } = buildSampleReports(groups, COMMON, 'Computing', 50);
    // First report: lowest displayOrder (index 0) from each group = "High"
    expect(reports[0].selections['g1'].code).toBe('High');
    expect(reports[0].selections['g2'].code).toBe('High');
  });

  it('tracks untested codes', () => {
    // 2 groups × 3 options each = 6 codes. With maxReports=1, some will be untested.
    const groups = [
      makeGroup('g1', 'Knowledge', ['H', 'M', 'L']),
      makeGroup('g2', 'Skills', ['H', 'M', 'L']),
    ];
    const { untestedItems } = buildSampleReports(groups, COMMON, 'Computing', 1);
    expect(untestedItems.length).toBeGreaterThan(0);
  });

  it('reports zero untested when all codes are covered', () => {
    // 1 group × 1 option = fully covered in 1 report
    const groups = [makeGroup('g1', 'Knowledge', ['OnlyOption'])];
    const { untestedItems } = buildSampleReports(groups, COMMON, 'Computing', 50);
    expect(untestedItems.length).toBe(0);
  });
});
