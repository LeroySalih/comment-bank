import { parseComment } from '@/lib/utils';

// ─── Shared types ─────────────────────────────────────────────────────────────

type MinimalOption = {
  id: string;
  code: string;
  text: string;
};

type MinimalGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommentOption: MinimalOption[];
};

type MinimalPupil = {
  firstName: string;
  lastName: string;
  gender: string;
};

type MinimalPupilCode = {
  groupId: string;
  code: string | null;
};

type MinimalAssignment = {
  id: string;
  Pupil: MinimalPupil;
  PupilCode: MinimalPupilCode[];
  eoyLevel?: string | null;
  targetLevel?: string | null;
  linkedData?: Record<string, string> | null;
};

type MinimalSubject = {
  studiedComment?: string | null;
  title?: string | null;
};

type MinimalClass = {
  year?: string | null;
};

type MinimalCommonGroup = {
  id: string;
  name: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommonCommentOption: MinimalOption[];
};

type MinimalCommonPupilCode = {
  commonGroupId: string;
  code: string | null;
};

// ─── Core assembly options ────────────────────────────────────────────────────

export interface CommentAssemblyOptions {
  /** Returns the selected raw option text for a subject group id, or '' */
  getSubjectText: (groupId: string) => string;
  /** Returns the selected raw option text for a common group id, or '' */
  getCommonText: (groupId: string) => string;
  /** Subject groups — name used for format template matching */
  subjectGroups: { id: string; name: string }[];
  /** Common groups — name used for format template matching */
  commonGroups: { id: string; name: string }[];
  /** Subject-studied blurb prepended to the SCG block */
  studiedComment?: string | null;
  /** Global format template (e.g. "<Academic>\n<SCG>\n<Behaviour>") */
  formatTemplate?: string;
  /** Subject-specific SCG ordering (e.g. "<Communication>\n<Make Skills>") */
  subjectFormat?: string | null;
}

/**
 * Single canonical comment assembly function.
 *
 * Resolves <GroupName> and <SCG> template tags using the provided getters,
 * returning the assembled text with pupil variable tags (<Name>, <He>, etc.)
 * still intact so the caller can substitute them with real or test values.
 */
export function assembleRawComment({
  getSubjectText,
  getCommonText,
  subjectGroups,
  commonGroups,
  studiedComment,
  formatTemplate,
  subjectFormat,
}: CommentAssemblyOptions): string {
  // Subject groups whose name matches a CCG name override that CCG slot
  const ccgNames = new Set(commonGroups.map(g => g.name));
  const overrideByName = new Map(
    subjectGroups.filter(g => ccgNames.has(g.name)).map(g => [g.name, g])
  );

  // Build the SCG block: pure subject groups (not CCG overrides), space-joined
  const buildScgBlock = (): string => {
    const parts: string[] = [];
    if (studiedComment) parts.push(studiedComment);

    const pureScg = subjectGroups.filter(g => !ccgNames.has(g.name));
    let ordered: { id: string; name: string }[];

    if (subjectFormat) {
      const tagPattern = /<([^>]+)>/g;
      ordered = [];
      let m: RegExpExecArray | null;
      while ((m = tagPattern.exec(subjectFormat)) !== null) {
        const g = pureScg.find(sg => sg.name === m![1]);
        if (g) ordered.push(g);
      }
    } else {
      ordered = pureScg;
    }

    const groupTexts = ordered.map(g => getSubjectText(g.id)).filter(Boolean);
    if (groupTexts.length > 0) parts.push(groupTexts.join(' '));
    return parts.join(' ');
  };

  if (formatTemplate) {
    let result = formatTemplate;

    // Replace each <CCGName> tag — subject override wins over CCG option
    for (const g of commonGroups) {
      const override = overrideByName.get(g.name);
      const text = override
        ? getSubjectText(override.id) || getCommonText(g.id)
        : getCommonText(g.id);
      result = result.replaceAll(`<${g.name}>`, text);
    }

    // Replace <SCG> with the assembled subject block
    result = result.replaceAll('<SCG>', buildScgBlock());

    // Strip any remaining unresolved custom tags, preserve standard variable tags
    const standardVars = new Set([
      'Name', 'He', 'he', 'She', 'she', 'His', 'his',
      'Her', 'her', 'Him', 'him', 'Subject', 'TargetLevel',
      'EoYLevel', 'Year', 'SCG',
      'he/she', 'his/her', 'him/her',
    ]);
    result = result.replace(/<([^>]+)>/g, (match, tag) =>
      standardVars.has(tag) ? match : ''
    );

    return result
      .split(/\n+/)
      .map(p => p.replace(/\s+/g, ' ').trim())
      .filter(Boolean)
      .join('\n\n');
  }

  // No format template: SCG block then remaining CCG groups
  const subjectGroupNames = new Set(subjectGroups.map(g => g.name));
  const paragraphs: string[] = [];

  const scgBlock = buildScgBlock();
  if (scgBlock) paragraphs.push(scgBlock);

  for (const g of commonGroups) {
    if (subjectGroupNames.has(g.name)) continue; // overridden by subject group
    const override = overrideByName.get(g.name);
    const text = override
      ? getSubjectText(override.id) || getCommonText(g.id)
      : getCommonText(g.id);
    if (text) paragraphs.push(text);
  }

  return paragraphs.join('\n\n');
}

// ─── Convenience wrapper for real assignment data ─────────────────────────────

/**
 * Generates the full comment string for an assignment.
 * Adapts assignment/group data into assembleRawComment, then substitutes
 * pupil variables via parseComment.
 */
export function generateComment(
  assignment: MinimalAssignment,
  subject: MinimalSubject,
  groups: MinimalGroup[],
  cls?: MinimalClass,
  commonGroups?: MinimalCommonGroup[],
  commonPupilCodes?: MinimalCommonPupilCode[],
  formatTemplate?: string,
  subjectFormat?: string | null
): string {
  const getSubjectText = (groupId: string): string => {
    const group = groups.find(g => g.id === groupId);
    if (!group) return '';
    if (group.isLinked && group.linkedField) {
      const code = assignment.linkedData?.[group.linkedField];
      if (!code) return '';
      return group.CommentOption.find(o => o.code === code)?.text || '';
    }
    const pc = assignment.PupilCode.find(c => c.groupId === groupId);
    if (!pc?.code) return '';
    return group.CommentOption.find(o => o.code === pc.code)?.text || '';
  };

  const getCommonText = (groupId: string): string => {
    const group = commonGroups?.find(g => g.id === groupId);
    if (!group) return '';
    if (group.isLinked && group.linkedField) {
      const code = assignment.linkedData?.[group.linkedField];
      if (!code) return '';
      return group.CommonCommentOption.find(o => o.code === code)?.text || '';
    }
    const cpc = commonPupilCodes?.find(c => c.commonGroupId === groupId);
    if (!cpc?.code) return '';
    return group.CommonCommentOption.find(o => o.code === cpc.code)?.text || '';
  };

  const raw = assembleRawComment({
    getSubjectText,
    getCommonText,
    subjectGroups: groups,
    commonGroups: commonGroups ?? [],
    studiedComment: subject.studiedComment,
    formatTemplate,
    subjectFormat,
  });

  return parseComment(
    raw,
    assignment.Pupil.firstName,
    assignment.Pupil.gender,
    subject.title || '',
    cls?.year,
    assignment.eoyLevel,
    assignment.targetLevel
  );
}
