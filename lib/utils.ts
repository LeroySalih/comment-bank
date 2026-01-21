export const parseComment = (commentText: string, firstName: string, gender: string) => {
    if (!commentText) return "";

    let comment = commentText;
    
    // Gender replacement logic
    const isMale = gender.toLowerCase() === "male";
    
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

    // Apply replacements
    // Note: The original code used replaceAll.
    Object.entries(replacements).forEach(([placeholder, value]) => {
        comment = comment.replaceAll(placeholder, value);
    });

    // Name replacement
    comment = comment.replaceAll('<Name>', firstName);

    // Capitalization fix for start of sentences if needed? 
    // The original code converted <He> -> he (lowercase). 
    // Wait, <He> usually starts a sentence. Original: commentText.replaceAll("<He>", "he").
    // If the template has "<He> is good", it becomes "he is good". That seems WRONG strictly speaking, but maybe the templates rely on previous sentence flow?
    // Let's check `main.js`:
    // comment = gender == "Male" ? commentText.replaceAll("<He>", "he") : commentText.replaceAll("<He>", "she")
    // Yes, it replaces <He> with lowercase he/she.
    // This implies <He> is NOT sentence start? Or the user's templates are weird?
    // Actually, look at line 27: `<He>` -> `he`.
    // Maybe `<He>` is just a token.
    // However, usually `<He>` implies Capitalized.
    // Let's stick EXACTLY to `main.js` logic to be safe.
    
    return comment;
}
