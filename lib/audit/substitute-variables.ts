const FIXED_VALUES: Record<string, string> = {
  '<Name>': 'Alex',
  '<he/she>': 'they',
  '<his/her>': 'their',
  '<him/her>': 'them',
  '<Year>': 'Year 10',
  '<EoYLevel>': '6',
  '<TargetLevel>': '7',
};

/**
 * Replaces all template variables in a comment with fixed audit test values.
 * <Subject> is replaced with the actual subject title from the DB.
 */
export function substituteVariables(text: string, subjectTitle: string): string {
  let result = text;
  for (const [variable, value] of Object.entries(FIXED_VALUES)) {
    result = result.replaceAll(variable, value);
  }
  result = result.replaceAll('<Subject>', subjectTitle);
  return result;
}
