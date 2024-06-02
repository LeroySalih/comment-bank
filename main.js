import xlsx from 'node-xlsx';
import PDFDocument from 'pdfkit'
import fs from 'fs'


const getCommentClassId = (classId) => {
    
    if (classId === undefined || classId === 'Class')
        return undefined;


    if (['13BS', '12BS', '11EC', '11CS', '10CS', '10DT', '10IT', '11IT'].includes(classId))
        return classId;

    if (['11BS1', '11BS2', '10BS1', '10BS2' ,  '10EC1', '10EC2'].includes(classId))
        return classId.substring(0, 4)

    return classId.substring(3, 5)
}

const parseComment = (commentText, firstName, gender) => {

    if (commentText == null) {
        return null;
    }
    let comment = gender == "Male" ? commentText.replaceAll("<He>", "he") : commentText.replaceAll("<He>", "she")
    comment = gender == "Male" ? comment.replaceAll("<he>", "he") : comment.replaceAll("<he>", "she")
    comment = gender == "Male" ? comment.replaceAll("<him>", "him") : comment.replaceAll("<him>", "her")
    comment = gender == "Male" ? comment.replaceAll("<his>", "his") : comment.replaceAll("<his>", "her")
    comment = comment.replaceAll('<Name>', firstName)
    return comment
}

const displayCommentsForPupil = (pupilDataItem, comments ) => {

    const [classId, familyName, firstName, gender, wpCode, thCode, psCode, oaCode, _] = pupilDataItem;
    
    // return {classId, familyName, givenName, wpCode, thCode, psCode, oaCode};

    const commentClassId = getCommentClassId(classId);

    if (commentClassId === undefined){
        return null;
    }

    const studiedKey = `${commentClassId}-STUDIED`; 
    const studiedComment = parseComment(comments[studiedKey], firstName, gender);

    const wpCodeKey = `${commentClassId}-${wpCode}`;
    const wpComment = parseComment(comments[wpCodeKey], firstName, gender);

    const thCodeKey = `${commentClassId}-${thCode}`;
    const thComment = parseComment(comments[thCodeKey], firstName, gender);

    const psCodeKey = `${commentClassId}-${psCode}`;
    const psComment = parseComment(comments[psCodeKey], firstName, gender);

    const oaCodeKey = `${commentClassId}-${oaCode}`;
    const oaComment = parseComment(comments[oaCodeKey], firstName, gender);



    return {classId, familyName, firstName, gender, studiedComment, wpComment, thComment, psComment, oaComment}

}

const createPDF = (comments) => {
// Create a document
const doc = new PDFDocument();
 
// Saving the pdf file in root directory.
doc.pipe(fs.createWriteStream('comments.pdf'));
 
for (const comment of comments) {

    if (!comment) {
        continue;
    }

    // Get a reference to the Outline root
    const { outline } = doc;

    // Add a top-level bookmark
    const top = outline.addItem(`${comment.classId} ${comment.firstName} ${comment.familyName} `);

    // console.log(comment);
    // Adding functionality
    doc
    .fontSize(12)
    .text(`${comment.classId} ${comment.firstName} ${comment.familyName} `, 100, 100)
    .moveDown()
    .text(comment.studiedComment)
    .moveDown()
    .text(`${comment.wpComment} ${comment.thComment} ${comment.psComment}`)
    .moveDown()
    .text(comment.oaComment)
    .moveDown()
    .addPage()
}
 
doc.end();
}

// Parse a file
const workSheetsFromFile = xlsx.parse(`./PupiL Lists 3.xlsx`);

const pupilSheet = workSheetsFromFile.find(sheet => sheet.name === 'Comment Banks');
const comments = pupilSheet.data.reduce((prev, cur) => { prev[`${cur[0]}-${cur[1]}`] = cur[2]; return prev}, {})

// console.log (comments);

const pupilData = workSheetsFromFile.find(sheet => sheet.name === 'Pupil Sheets');
//console.log(pupilData)

const pupilComments = pupilData.data.map(pd => displayCommentsForPupil(pd, comments));
console.log(pupilComments);

createPDF(pupilComments);


