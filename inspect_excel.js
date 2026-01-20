const xlsx = require('xlsx');
const file = xlsx.readFile('../Pupil Lists 24-25.xlsx');
const sheet = file.Sheets['Comment Banks'];
const data = xlsx.utils.sheet_to_json(sheet, {header: 1});
console.log("First 10 rows of associated data:");
console.log(JSON.stringify(data.slice(0, 10), null, 2));

// Also check a class sheet to see codes
const classSheetName = file.SheetNames.find(n => n !== 'Comment Banks' && n !== 'Pupil Sheets');
if (classSheetName) {
    console.log(`\nClass Sheet: ${classSheetName}`);
    const classSheet = file.Sheets[classSheetName];
    const classData = xlsx.utils.sheet_to_json(classSheet, {header: 1});
    console.log(JSON.stringify(classData.slice(0, 5), null, 2));
}
