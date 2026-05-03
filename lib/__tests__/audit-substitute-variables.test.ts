import { describe, it, expect } from 'vitest';
import { substituteVariables } from '../audit/substitute-variables';

describe('substituteVariables', () => {
  it('replaces <Name> with Alex', () => {
    expect(substituteVariables('<Name> is a great student.', 'Computing')).toBe(
      'Alex is a great student.'
    );
  });

  it('replaces <he/she> with they', () => {
    expect(substituteVariables('<he/she> works hard.', 'Computing')).toBe(
      'they works hard.'
    );
  });

  it('replaces <his/her> with their', () => {
    expect(substituteVariables('<his/her> work is excellent.', 'Computing')).toBe(
      'their work is excellent.'
    );
  });

  it('replaces <him/her> with them', () => {
    expect(substituteVariables('I encourage <him/her>.', 'Computing')).toBe(
      'I encourage them.'
    );
  });

  it('replaces <Subject> with the provided subject title', () => {
    expect(substituteVariables('<Name> enjoys <Subject>.', 'French')).toBe(
      'Alex enjoys French.'
    );
  });

  it('replaces <Year> with Year 10', () => {
    expect(substituteVariables('In <Year>, <Name> excelled.', 'Maths')).toBe(
      'In Year 10, Alex excelled.'
    );
  });

  it('replaces <EoYLevel> with 6', () => {
    expect(substituteVariables('Achieved <EoYLevel>.', 'Maths')).toBe(
      'Achieved 6.'
    );
  });

  it('replaces <TargetLevel> with 7', () => {
    expect(substituteVariables('Target is <TargetLevel>.', 'Maths')).toBe(
      'Target is 7.'
    );
  });

  it('replaces all occurrences of the same variable', () => {
    expect(substituteVariables('<Name> and <Name> again.', 'Computing')).toBe(
      'Alex and Alex again.'
    );
  });

  it('returns text unchanged when no variables present', () => {
    expect(substituteVariables('No variables here.', 'Computing')).toBe(
      'No variables here.'
    );
  });
});
