import { substituteVariables } from './substitute-variables';
import { assembleRawComment } from '@/lib/comment-utils';
import type { SampledReport } from './types';

type Option = { id: string; code: string; text: string; displayOrder: number };
type Group = { id: string; name: string; title: string; isLinked: boolean; options: Option[] };

type BuildResult = {
  reports: SampledReport[];
  untestedItems: { code: string; groupName: string }[];
};

/** Sort options by displayOrder ascending; index 0 = "High" */
function sortedOptions(options: Option[]): Option[] {
  return [...options].sort((a, b) => a.displayOrder - b.displayOrder);
}

function pickOption(options: Option[], tier: 'high' | 'medium' | 'low'): Option {
  const sorted = sortedOptions(options);
  if (tier === 'high') return sorted[0];
  if (tier === 'low') return sorted[sorted.length - 1];
  return sorted[Math.round((sorted.length - 1) / 2)];
}

export function buildSampleReports(
  subjectGroups: Group[],
  commonGroups: Group[],
  subjectTitle: string,
  formatTemplate = '',
  subjectFormat: string | null = null
): BuildResult {
  // Filter out linked groups; groups must have at least one option
  const activeGroups = subjectGroups.filter(g => !g.isLinked && g.options.length > 0);

  if (activeGroups.length === 0) {
    return { reports: [], untestedItems: [] };
  }

  const reports: SampledReport[] = [];
  const seenCodes = new Set<string>(); // `${groupId}:${code}`

  function addReport(
    subjectSelections: Record<string, Option>,
    commonSelectionsForReport: Record<string, Option>,
    label: string,
  ): void {
    const index = reports.length;
    const selectionsForReport: SampledReport['selections'] = {};

    for (const [groupId, opt] of Object.entries(subjectSelections)) {
      const group = activeGroups.find(g => g.id === groupId)!;
      selectionsForReport[groupId] = { code: opt.code, text: opt.text, groupTitle: group.title };
      seenCodes.add(`${groupId}:${opt.code}`);
    }
    for (const [groupId, opt] of Object.entries(commonSelectionsForReport)) {
      const cg = commonGroups.find(g => g.id === groupId)!;
      selectionsForReport[groupId] = { code: opt.code, text: opt.text, groupTitle: cg.title };
    }

    // Use the single canonical assembly function, then substitute audit variables
    const raw = assembleRawComment({
      getSubjectText: (groupId) => {
        const opt = subjectSelections[groupId];
        return opt?.text ?? '';
      },
      getCommonText: (groupId) => {
        const opt = commonSelectionsForReport[groupId];
        return opt?.text ?? '';
      },
      subjectGroups: activeGroups,
      commonGroups,
      formatTemplate,
      subjectFormat,
    });

    const assembledText = substituteVariables(raw, subjectTitle);
    reports.push({ reportIndex: index, label, selections: selectionsForReport, assembledText });
  }

  // Build common selections for a given tier
  function buildCommonSelections(tier: 'high' | 'medium' | 'low'): Record<string, Option> {
    const sel: Record<string, Option> = {};
    for (const cg of commonGroups) {
      if (!cg.isLinked && cg.options.length > 0) {
        sel[cg.id] = pickOption(cg.options, tier);
      }
    }
    return sel;
  }

  // ── Strategy 1: All High ────────────────────────────────────────────────
  {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'high');
    addReport(sel, buildCommonSelections('high'), 'All High');
  }

  // ── Strategy 2: All Medium ──────────────────────────────────────────────
  {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'medium');
    addReport(sel, buildCommonSelections('medium'), 'All Medium');
  }

  // ── Strategy 3: All Low ─────────────────────────────────────────────────
  {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'low');
    addReport(sel, buildCommonSelections('low'), 'All Low');
  }

  // ── Strategy 4: Coverage fill ─────────────────────────────────────────────
  // Adds extra rows only when a group has 2 or 4 options, leaving some
  // options uncovered by the three base rows above.
  let coverageCount = 0;
  while (true) {
    const unseenOptions: { groupId: string; opt: Option }[] = [];
    for (const g of activeGroups) {
      for (const opt of g.options) {
        if (!seenCodes.has(`${g.id}:${opt.code}`)) {
          unseenOptions.push({ groupId: g.id, opt });
        }
      }
    }
    if (unseenOptions.length === 0) break;

    coverageCount++;
    const sel: Record<string, Option> = {};
    const coveredGroups = new Set<string>();
    for (const { groupId, opt } of unseenOptions) {
      if (!coveredGroups.has(groupId)) {
        sel[groupId] = opt;
        coveredGroups.add(groupId);
      }
    }
    for (const g of activeGroups) {
      if (!sel[g.id]) sel[g.id] = pickOption(g.options, 'high');
    }
    addReport(sel, buildCommonSelections('high'), `Coverage Fill ${coverageCount}`);
  }

  // ── Untested codes ───────────────────────────────────────────────────────
  const untestedItems: { code: string; groupName: string }[] = [];
  for (const g of activeGroups) {
    for (const opt of g.options) {
      if (!seenCodes.has(`${g.id}:${opt.code}`)) {
        untestedItems.push({ code: opt.code, groupName: g.title });
      }
    }
  }

  return { reports, untestedItems };
}
