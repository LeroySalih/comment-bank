
import { generateComment } from './lib/comment-utils';

// Mock data
const mockStudent = {
  id: '1',
  firstName: 'John',
  lastName: 'Doe',
  gender: 'Male',
  wpCode: 'H',
  thCode: 'M',
  psCode: 'L',
  oaCode: 'H',
  eoyLevel: '5A',
  targetLevel: '6B'
};

const mockCourse = {
  studiedComment: 'This term <Subject> students studied Shakespeare.',
  subject: 'English'
};

const mockClass = {
  year: '9'
};

const mockGroups = [
  {
    id: 'g1',
    name: 'WP',
    options: [{ id: 'o1', code: 'H', text: '<He> has worked hard in Year <Year>.' }]
  },
  {
    id: 'g2',
    name: 'TH',
    options: [{ id: 'o2', code: 'M', text: '<He> shows moderate understanding. Target: <TargetLevel>.' }]
  },
  {
    id: 'g3',
    name: 'PS',
    options: [{ id: 'o3', code: 'L', text: '<His> practical skills need improvement. EoY: <EoYLevel>.' }]
  },
  {
    id: 'g4',
    name: 'OA',
    options: [{ id: 'o4', code: 'H', text: 'Overall excellent.' }]
  }
];

// Expected Output:
// This term English students studied Shakespeare.
//
// He has worked hard in Year 9. he shows moderate understanding. Target: 6B. his practical skills need improvement. EoY: 5A.
//
// Overall excellent.

console.log("Running Test...");
const result = generateComment(mockStudent, mockCourse, mockGroups, mockClass);
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
