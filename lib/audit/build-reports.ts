import { substituteVariables } from './substitute-variables';
import type { SampledReport } from './types';

type Option = { id: string; code: string; text: string; displayOrder: number };
type Group = { id: string; title: string; isLinked: boolean; options: Option[] };

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
  return sorted[Math.floor((sorted.length - 1) / 2)];
}

function assembleText(
  subjectSelections: Record<string, Option>,
  commonSelections: Record<string, Option>,
  subjectTitle: string
): string {
  const parts: string[] = [];
  for (const opt of Object.values(subjectSelections)) {
    parts.push(substituteVariables(opt.text, subjectTitle));
  }
  for (const opt of Object.values(commonSelections)) {
    parts.push(substituteVariables(opt.text, subjectTitle));
  }
  return parts.join(' ');
}

export function buildSampleReports(
  subjectGroups: Group[],
  commonGroups: Group[],
  subjectTitle: string,
  maxReports = 50
): BuildResult {
  // Filter out linked groups; groups must have at least one option
  const activeGroups = subjectGroups.filter(g => !g.isLinked && g.options.length > 0);

  // Fixed common selections: first option by displayOrder from each common group
  const commonSelections: Record<string, Option> = {};
  for (const cg of commonGroups) {
    if (!cg.isLinked && cg.options.length > 0) {
      commonSelections[cg.id] = sortedOptions(cg.options)[0];
    }
  }

  const reports: SampledReport[] = [];

  // Track which (groupId, code) combos have been seen
  const seenCodes = new Set<string>(); // `${groupId}:${code}`

  function addReport(subjectSelections: Record<string, Option>): boolean {
    if (reports.length >= maxReports) return false;
    const index = reports.length;
    const selectionsForReport: SampledReport['selections'] = {};
    for (const [groupId, opt] of Object.entries(subjectSelections)) {
      const group = activeGroups.find(g => g.id === groupId)!;
      selectionsForReport[groupId] = {
        code: opt.code,
        text: opt.text,
        groupTitle: group.title,
      };
      seenCodes.add(`${groupId}:${opt.code}`);
    }
    // Also record common selections
    for (const [groupId, opt] of Object.entries(commonSelections)) {
      const cg = commonGroups.find(g => g.id === groupId)!;
      selectionsForReport[groupId] = {
        code: opt.code,
        text: opt.text,
        groupTitle: cg.title,
      };
    }
    const assembledText = assembleText(subjectSelections, commonSelections, subjectTitle);
    reports.push({ reportIndex: index, selections: selectionsForReport, assembledText });
    return true;
  }

  // ── Strategy 1: All High ────────────────────────────────────────────────
  {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'high');
    addReport(sel);
  }

  // ── Strategy 2: All Medium ──────────────────────────────────────────────
  if (reports.length < maxReports) {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'medium');
    addReport(sel);
  }

  // ── Strategy 3: All Low ─────────────────────────────────────────────────
  if (reports.length < maxReports) {
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) sel[g.id] = pickOption(g.options, 'low');
    addReport(sel);
  }

  // ── Strategy 4: Mostly High (one group rotated to Medium) ───────────────
  for (const rotateGroup of activeGroups) {
    if (reports.length >= maxReports) break;
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) {
      sel[g.id] = g.id === rotateGroup.id
        ? pickOption(g.options, 'medium')
        : pickOption(g.options, 'high');
    }
    addReport(sel);
  }

  // ── Strategy 5: Mostly Low (one group rotated to High) ──────────────────
  for (const rotateGroup of activeGroups) {
    if (reports.length >= maxReports) break;
    const sel: Record<string, Option> = {};
    for (const g of activeGroups) {
      sel[g.id] = g.id === rotateGroup.id
        ? pickOption(g.options, 'high')
        : pickOption(g.options, 'low');
    }
    addReport(sel);
  }

  // ── Strategy 6: Coverage fill ────────────────────────────────────────────
  while (reports.length < maxReports) {
    const unseenOptions: { groupId: string; opt: Option }[] = [];
    for (const g of activeGroups) {
      for (const opt of g.options) {
        if (!seenCodes.has(`${g.id}:${opt.code}`)) {
          unseenOptions.push({ groupId: g.id, opt });
        }
      }
    }
    if (unseenOptions.length === 0) break;

    // Build one report that covers as many unseen options as possible
    const sel: Record<string, Option> = {};
    const coveredGroups = new Set<string>();

    // First pass: pick unseen options
    for (const { groupId, opt } of unseenOptions) {
      if (!coveredGroups.has(groupId)) {
        sel[groupId] = opt;
        coveredGroups.add(groupId);
      }
    }
    // Fill remaining groups with their "high" option
    for (const g of activeGroups) {
      if (!sel[g.id]) sel[g.id] = pickOption(g.options, 'high');
    }
    addReport(sel);
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
