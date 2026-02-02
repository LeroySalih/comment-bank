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
};

type MinimalSubject = {
  studiedComment?: string | null;
  title?: string | null;
};

type MinimalClass = {
    year?: string | null;
}

/**
 * Generates the full comment string for an assignment based on PupilCodes and Subject configuration.
 */
export function generateComment(
  assignment: MinimalAssignment,
  subject: MinimalSubject,
  groups: MinimalGroup[],
  cls?: MinimalClass
): string {
    // Helper to get selected option text for a group
    const getOptionTextForGroup = (group: MinimalGroup): string => {
        const pc = assignment.PupilCode?.find(c => c.groupId === group.id);
        if (!pc || !pc.code) return "";
        const option = group.CommentOption?.find(o => o.code === pc.code);
        return option?.text || "";
    };

    const parts: string[] = [];

    // Add studied comment if present
    if (subject.studiedComment) {
        parts.push(subject.studiedComment);
    }

    // Iterate through all groups in order and collect their selected texts
    const groupTexts: string[] = [];
    for (const group of groups) {
        const text = getOptionTextForGroup(group);
        if (text) {
            groupTexts.push(text);
        }
    }

    // Join all group texts with spaces
    if (groupTexts.length > 0) {
        parts.push(groupTexts.join(" "));
    }

    // Combine all parts with paragraph breaks
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
