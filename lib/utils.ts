export const parseComment = (
    commentText: string, 
    firstName: string, 
    gender: string,
    // New optional fields
    subject?: string | null,
    year?: string | null,
    eoyLevel?: string | null,
    targetLevel?: string | null
) => {
    if (!commentText) return "";

    let comment = commentText;
    
    // Gender replacement logic
    const g = gender.toLowerCase();
    const isMale = g === "male" || g === "m";
    
    const replacements = isMale ? {
        "<He>": "He",
        "<he>": "he",
        "<Him>": "Him", // Rarely used at start, but consistent
        "<him>": "him",
        "<His>": "His",
        "<his>": "his",
        "<She>": "He",
        "<she>": "he",
        "<Her>": "Him",
    } : {
        "<He>": "She",
        "<he>": "she",
        "<Him>": "Her",
        "<him>": "her",
        "<His>": "Her", // "His book" -> "Her book" (possessive)
        "<his>": "her"
    };

    // Apply gender replacements
    Object.entries(replacements).forEach(([placeholder, value]) => {
        comment = comment.replaceAll(placeholder, value);
    });

    // Name replacement
    comment = comment.replaceAll('<Name>', firstName);

    // New Variable Replacements
    if (subject) comment = comment.replaceAll('<Subject>', subject);
    if (year) comment = comment.replaceAll('<Year>', year);
    if (eoyLevel) comment = comment.replaceAll('<EoYLevel>', eoyLevel);
    if (targetLevel) comment = comment.replaceAll('<TargetLevel>', targetLevel);

    return comment;
}

export const countWords = (text: string) => {
    if (!text) return 0;
    return text.trim().split(/\s+/).filter(word => word.length > 0).length;
}
