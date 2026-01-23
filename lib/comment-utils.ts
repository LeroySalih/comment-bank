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
  options: MinimalOption[];
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
  pupil: MinimalPupil;
  codes: MinimalPupilCode[];
  eoyLevel?: string | null;
  targetLevel?: string | null;
};

type MinimalSubject = {
  studiedComment?: string | null;
  subject?: string | null;
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
    // Helper to find option text by group name
    const getOptionText = (groupName: string): string => {
        const group = groups.find(g => g.name === groupName);
        if (!group) return "";
        const pc = assignment.codes.find(c => c.groupId === group.id);
        if (!pc || !pc.code) return "";
        const option = group.options.find(o => o.code === pc.code);
        return option?.text || "";
    };

    const studied = subject.studiedComment || "";
    
    // Fetch texts for standard groups
    const wp = getOptionText("WP");
    const th = getOptionText("TH");
    const ps = getOptionText("PS");
    const oa = getOptionText("OA");

    let combined = "";
    if (studied) combined += studied + "\n\n";
    
    // Middle block: WP, TH, PS joined by spaces
    const middleBlock = [wp, th, ps].filter(Boolean).join(" ");
    if (middleBlock) combined += middleBlock + "\n\n";

    // OA is separate paragraph usually
    if (oa) combined += oa;

    return parseComment(
        combined, 
        assignment.pupil.firstName, 
        assignment.pupil.gender,
        subject.subject,
        cls?.year,
        assignment.eoyLevel,
        assignment.targetLevel
    );
}
