import { generateComment } from './lib/comment-utils';

// Mock data
const mockAssignment: any = {
  id: '1',
  Pupil: {
    admissionNumber: '123456',
    firstName: 'John',
    lastName: 'Doe',
    gender: 'Male',
  },
  PupilCode: [
    { groupId: 'g1', code: 'H' },
    { groupId: 'g2', code: 'M' },
    { groupId: 'g3', code: 'L' },
    { groupId: 'g4', code: 'H' }
  ],
  eoyLevel: '5A',
  targetLevel: '6B'
};

const mockSubject = {
  studiedComment: 'This term <Subject> students studied Shakespeare.',
  title: 'English'
};

const mockClass = {
  year: '9'
};

const mockGroups = [
  {
    id: 'g1',
    name: 'WP',
    CommentOption: [{ id: 'o1', code: 'H', text: '<He> has worked hard in Year <Year>.' }]
  },
  {
    id: 'g2',
    name: 'TH',
    CommentOption: [{ id: 'o2', code: 'M', text: '<He> shows moderate understanding. Target: <TargetLevel>.' }]
  },
  {
    id: 'g3',
    name: 'PS',
    CommentOption: [{ id: 'o3', code: 'L', text: '<His> practical skills need improvement. EoY: <EoYLevel>.' }]
  },
  {
    id: 'g4',
    name: 'OA',
    CommentOption: [{ id: 'o4', code: 'H', text: 'Overall excellent.' }]
  }
];

console.log("Running Test...");
const result = generateComment(mockAssignment, mockSubject, mockGroups, mockClass);
console.log("Result:");
console.log(result);

if (result.includes("This term English students studied Shakespeare.") &&
    result.includes("Overall excellent.") &&
    result.includes("He has worked hard in Year 9") &&
    result.includes("Target: 6B") &&
    result.includes("EoY: 5A")) {
    console.log("✅ simple content checks passed");
} else {
    console.log("❌ content checks failed");
}
