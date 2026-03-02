import { parseComment } from '@/lib/utils';

// Define types needed for the generator
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
}

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

/**
 * Generates the full comment string for an assignment using the format template system.
 *
 * The format template uses <GroupName> tags for CCG groups and <SCG> for subject comment groups.
 * For <SCG>, the subject's commentFormat defines the order of groups to include.
 * <Subject> remains a standard variable replaced with the subject title.
 *
 * When no format template is provided, falls back to joining all group texts.
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
    // Helper to get selected option text for a subject-specific group
    const getOptionTextForGroup = (group: MinimalGroup): string => {
        if (group.isLinked && group.linkedField) {
            const linkedData = assignment.linkedData as Record<string, string> | null | undefined;
            const code = linkedData?.[group.linkedField];
            if (!code) return "";
            const option = group.CommentOption?.find(o => o.code === code);
            return option?.text || "";
        }
        const pc = assignment.PupilCode?.find(c => c.groupId === group.id);
        if (!pc || !pc.code) return "";
        const option = group.CommentOption?.find(o => o.code === pc.code);
        return option?.text || "";
    };

    // Helper to get selected option text for a common group
    const getCommonOptionText = (group: MinimalCommonGroup): string => {
        if (group.isLinked && group.linkedField) {
            const linkedData = assignment.linkedData as Record<string, string> | null | undefined;
            const code = linkedData?.[group.linkedField];
            if (!code) return "";
            const option = group.CommonCommentOption?.find(o => o.code === code);
            return option?.text || "";
        }
        const pc = commonPupilCodes?.find(c => c.commonGroupId === group.id);
        if (!pc || !pc.code) return "";
        const option = group.CommonCommentOption?.find(o => o.code === pc.code);
        return option?.text || "";
    };

    // Identify subject groups that override a CCG group (name matches a CCG group name)
    const ccgGroupNames = new Set((commonGroups ?? []).map(g => g.name));
    const overrideGroups = groups.filter(g => ccgGroupNames.has(g.name));
    const overrideGroupsByName = new Map(overrideGroups.map(g => [g.name, g]));

    // Build the subject content
    const buildSubjectContent = (): string => {
        const parts: string[] = [];
        if (subject.studiedComment) {
            parts.push(subject.studiedComment);
        }

        // Pure SCG groups only — exclude any group that overrides a CCG variable
        const pureScgGroups = groups.filter(g => !ccgGroupNames.has(g.name));

        let orderedGroups: MinimalGroup[];
        if (subjectFormat) {
            // Use the subject's comment format to order groups
            const codes = subjectFormat.split(/\s+/).filter(Boolean);
            orderedGroups = [];
            for (const code of codes) {
                const group = pureScgGroups.find(g => g.name === code);
                if (group) orderedGroups.push(group);
            }
        } else {
            // Default: all groups in display order
            orderedGroups = pureScgGroups;
        }

        const groupTexts: string[] = [];
        for (const group of orderedGroups) {
            const text = getOptionTextForGroup(group);
            if (text) groupTexts.push(text);
        }
        if (groupTexts.length > 0) parts.push(groupTexts.join(" "));
        return parts.join(" ");
    };

    // If we have a format template, use it
    if (formatTemplate && commonGroups && commonGroups.length > 0) {
        let result = formatTemplate;

        // Replace each CCG group tag — use subject override if available, fall back to CCG
        for (const group of commonGroups) {
            const overrideGroup = overrideGroupsByName.get(group.name);
            let text: string;
            if (overrideGroup) {
                text = getOptionTextForGroup(overrideGroup);
                if (!text) text = getCommonOptionText(group); // fallback to CCG
            } else {
                text = getCommonOptionText(group);
            }
            result = result.replaceAll(`<${group.name}>`, text);
        }

        // Replace <SCG> tag with subject comment group content
        const subjectContent = buildSubjectContent();
        result = result.replaceAll('<SCG>', subjectContent);

        // Clean up any unreplaced custom tags (but not standard variable tags)
        // Only remove tags that are not standard variables
        const standardVars = ['Name', 'He', 'he', 'She', 'she', 'His', 'his', 'Her', 'her', 'Him', 'him', 'Subject', 'TargetLevel', 'EoYLevel', 'Year', 'SCG'];
        result = result.replace(/<([^>]+)>/g, (match, tagName) => {
            if (standardVars.includes(tagName)) return match;
            return '';
        });
        result = result.split(/\n+/).map(p => p.replace(/\s+/g, ' ').trim()).filter(p => p).join('\n\n');

        return parseComment(
            result,
            assignment.Pupil.firstName,
            assignment.Pupil.gender,
            subject.title || '',
            cls?.year,
            assignment.eoyLevel,
            assignment.targetLevel
        );
    }

    // No format template: legacy behavior - just join all groups
    if (!commonGroups || commonGroups.length === 0) {
        const subjectContent = buildSubjectContent();
        return parseComment(
            subjectContent,
            assignment.Pupil.firstName,
            assignment.Pupil.gender,
            subject.title || '',
            cls?.year,
            assignment.eoyLevel,
            assignment.targetLevel
        );
    }

    // Has common groups but no template: join common groups then subject groups
    const paragraphs: string[] = [];

    const commonTexts = commonGroups.map(g => {
        const overrideGroup = overrideGroupsByName.get(g.name);
        if (overrideGroup) {
            const text = getOptionTextForGroup(overrideGroup);
            return text || getCommonOptionText(g);
        }
        return getCommonOptionText(g);
    }).filter(Boolean);
    if (commonTexts.length > 0) paragraphs.push(commonTexts.join(" "));

    const subjectContent = buildSubjectContent();
    if (subjectContent) paragraphs.push(subjectContent);

    const combined = paragraphs.join("\n\n");

    return parseComment(
        combined,
        assignment.Pupil.firstName,
        assignment.Pupil.gender,
        subject.title || '',
        cls?.year,
        assignment.eoyLevel,
        assignment.targetLevel
    );
}
