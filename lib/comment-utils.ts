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

type MinimalStudent = {
  id: string;
  firstName: string;
  lastName: string;
  gender: string;
  // Codes
  wpCode?: string | null;
  thCode?: string | null;
  psCode?: string | null;
  oaCode?: string | null;
};

type MinimalCourse = {
  studiedComment?: string | null;
};

/**
 * Generates the full comment string for a student based on their codes and course configuration.
 * Replicates the logic from CommentEditor.tsx
 */
export function generateComment(
  student: MinimalStudent,
  course: MinimalCourse,
  groups: MinimalGroup[]
): string {
    // Helper to find option text by group name and code
    const getOptionText = (groupName: string, code: string | null | undefined): string => {
        if (!code) return "";
        const group = groups.find(g => g.name === groupName);
        const option = group?.options.find(o => o.code === code);
        return option?.text || "";
    };

    const parts: string[] = [];

    // 1. Studied Comment (Course Intro)
    const studied = course.studiedComment || "";
    
    // 2. Fetch texts for standard groups
    const wp = getOptionText("WP", student.wpCode);
    const th = getOptionText("TH", student.thCode);
    const ps = getOptionText("PS", student.psCode);
    const oa = getOptionText("OA", student.oaCode);

    let combined = "";
    if (studied) combined += studied + "\n\n";
    
    // Middle block: WP, TH, PS joined by spaces
    const middleBlock = [wp, th, ps].filter(Boolean).join(" ");
    if (middleBlock) combined += middleBlock + "\n\n";

    // OA is separate paragraph usually
    if (oa) combined += oa;

    // Handle unknown groups (append at end if any exist in the student object? 
    // The original CommentEditor handled 'otherGroups' from `groups` prop based on `selections` state.
    // In this static context, we only have the student's codes. 
    // The current Student model in schema only has wpCode, thCode, psCode, oaCode. 
    // There are no 'other' codes stored on the student model yet (based on schema and types).
    // So for now, we only handle the standard 4. 
    // If dynamic groups are added later to the schema, this needs update.
    
    return parseComment(combined, student.firstName, student.gender);
}
