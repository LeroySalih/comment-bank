
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
  oaCode: 'H'
};

const mockCourse = {
  studiedComment: 'This term we studied Shakespeare.'
};

const mockGroups = [
  {
    id: 'g1',
    name: 'WP',
    options: [{ id: 'o1', code: 'H', text: '<He> has worked hard.' }]
  },
  {
    id: 'g2',
    name: 'TH',
    options: [{ id: 'o2', code: 'M', text: '<He> shows moderate understanding.' }]
  },
  {
    id: 'g3',
    name: 'PS',
    options: [{ id: 'o3', code: 'L', text: '<His> practical skills need improvement.' }]
  },
  {
    id: 'g4',
    name: 'OA',
    options: [{ id: 'o4', code: 'H', text: 'Overall excellent.' }]
  }
];

// Expected Output:
// This term we studied Shakespeare.
//
// He has worked hard. he shows moderate understanding. his practical skills need improvement.
//
// Overall excellent.

console.log("Running Test...");
const result = generateComment(mockStudent, mockCourse, mockGroups);
console.log("Result:");
console.log(result);

const expectedMiddle = "He has worked hard. he shows moderate understanding. his practical skills need improvement.";

if (result.includes("This term we studied Shakespeare.") && 
    result.includes("Overall excellent.") &&
    result.includes("He has worked hard")) {
    console.log("✅ simple content checks passed");
} else {
    console.log("❌ content checks failed");
}
