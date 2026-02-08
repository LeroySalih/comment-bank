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
  paragraphPosition: string;
  isLinked?: boolean;
  linkedField?: string | null;
  CommonCommentOption: MinimalOption[];
};

type MinimalCommonPupilCode = {
  commonGroupId: string;
  code: string | null;
};

/**
 * Generates the full comment string for an assignment based on PupilCodes and Subject configuration.
 * When commonGroups is provided, produces a 4-paragraph layout:
 *   P1: CCGs with paragraphPosition "p1"
 *   P2: CCGs with paragraphPosition "p2" — substituted into wrapperTemplate
 *   P3: Subject studiedComment + subject-specific group texts
 *   P4: CCGs with paragraphPosition "p4"
 * When commonGroups is undefined, only P3 is generated (backward compatible).
 */
export function generateComment(
  assignment: MinimalAssignment,
  subject: MinimalSubject,
  groups: MinimalGroup[],
  cls?: MinimalClass,
  commonGroups?: MinimalCommonGroup[],
  commonPupilCodes?: MinimalCommonPupilCode[],
  wrapperTemplate?: string
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

    // If no common groups provided, use legacy single-paragraph behavior
    if (!commonGroups) {
        const parts: string[] = [];

        if (subject.studiedComment) {
            parts.push(subject.studiedComment);
        }

        const groupTexts: string[] = [];
        for (const group of groups) {
            const text = getOptionTextForGroup(group);
            if (text) {
                groupTexts.push(text);
            }
        }

        if (groupTexts.length > 0) {
            parts.push(groupTexts.join(" "));
        }

        const combined = parts.join("\n\n");

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

    // 4-paragraph layout
    const paragraphs: string[] = [];

    // P1: CCGs with paragraphPosition === "p1"
    const p1Groups = commonGroups.filter(g => g.paragraphPosition === 'p1');
    const p1Texts = p1Groups.map(g => getCommonOptionText(g)).filter(Boolean);
    if (p1Texts.length > 0) {
        paragraphs.push(p1Texts.join(" "));
    }

    // P2: CCGs with paragraphPosition === "p2" — substitute into wrapper template
    const p2Groups = commonGroups.filter(g => g.paragraphPosition === 'p2');
    if (p2Groups.length > 0 && wrapperTemplate) {
        let p2Text = wrapperTemplate;
        for (const group of p2Groups) {
            const text = getCommonOptionText(group);
            // Replace <GroupName> placeholder with the resolved text
            p2Text = p2Text.replaceAll(`<${group.name}>`, text);
        }
        // Clean up any unreplaced placeholders (groups with no selection)
        p2Text = p2Text.replace(/<[^>]+>/g, '').replace(/\s{2,}/g, ' ').trim();
        if (p2Text) {
            paragraphs.push(p2Text);
        }
    } else if (p2Groups.length > 0) {
        // Fallback: just join texts if no wrapper template
        const p2Texts = p2Groups.map(g => getCommonOptionText(g)).filter(Boolean);
        if (p2Texts.length > 0) {
            paragraphs.push(p2Texts.join(" "));
        }
    }

    // P3: Subject studiedComment + subject-specific group texts
    const p3Parts: string[] = [];
    if (subject.studiedComment) {
        p3Parts.push(subject.studiedComment);
    }
    const groupTexts: string[] = [];
    for (const group of groups) {
        const text = getOptionTextForGroup(group);
        if (text) {
            groupTexts.push(text);
        }
    }
    if (groupTexts.length > 0) {
        p3Parts.push(groupTexts.join(" "));
    }
    if (p3Parts.length > 0) {
        paragraphs.push(p3Parts.join(" "));
    }

    // P4: CCGs with paragraphPosition === "p4"
    const p4Groups = commonGroups.filter(g => g.paragraphPosition === 'p4');
    const p4Texts = p4Groups.map(g => getCommonOptionText(g)).filter(Boolean);
    if (p4Texts.length > 0) {
        paragraphs.push(p4Texts.join(" "));
    }

    // Join non-empty paragraphs with double newlines
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
