import { PrismaClient } from '@prisma/client';
import xlsx from 'xlsx';
import fs from 'fs';
import path from 'path';
import { hash } from 'bcryptjs';

const prisma = new PrismaClient();
const EXCEL_PATH = '../Pupil Lists 24-25.xlsx';

async function main() {
  const absolutePath = path.resolve(EXCEL_PATH);
  if (!fs.existsSync(absolutePath)) {
    console.error(`File not found: ${absolutePath}`);
    process.exit(1);
  }

  // Seed Admin User
  const password = await hash('password', 10);
  await prisma.user.upsert({
    where: { username: 'admin' },
    update: {},
    create: {
      username: 'admin',
      password
    }
  });
  console.log('Seeded User: admin / password');

  const file = xlsx.readFile(absolutePath);
  
  // 1. Parse Comment Banks
  const bankSheet = file.Sheets['Comment Banks'];
  const bankData = xlsx.utils.sheet_to_json(bankSheet, { header: 1 }) as string[][];

  // Map: CourseName -> { studied: string, groups: { [groupName]: { options: { [code]: string } } } }
  const courses: Record<string, { studied?: string, groups: Record<string, Record<string, string>> }> = {};

  // Skip header if exists? Dump showed no header, just data. Row 0 was data.
  // [ "11CS", "STUDIED", "..." ]
  
  for (const row of bankData) {
    if (!row || row.length < 3) continue;
    const courseName = row[0]; // e.g. "11CS"
    const key = row[1];        // e.g. "STUDIED", "WP-H"
    const text = row[2];

    if (!courses[courseName]) {
      courses[courseName] = { groups: {} };
    }

    if (key === 'STUDIED') {
      courses[courseName].studied = text;
    } else {
      // Parse Key: "WP-H" -> Group "WP", Code "H"
      const parts = key.split('-');
      if (parts.length === 2) {
        const groupName = parts[0];
        const code = parts[1];
        
        if (!courses[courseName].groups[groupName]) {
          courses[courseName].groups[groupName] = {};
        }
        courses[courseName].groups[groupName][code] = text;
      }
    }
  }

  // Insert Courses and Comments
  for (const [courseName, data] of Object.entries(courses)) {
    console.log(`Seeding Course: ${courseName}`);
    
    const course = await prisma.course.upsert({
      where: { name: courseName },
      update: { studiedComment: data.studied },
      create: { 
        name: courseName,
        studiedComment: data.studied 
      }
    });

    for (const [groupName, options] of Object.entries(data.groups)) {
      const group = await prisma.commentGroup.upsert({
        where: { courseId_name: { courseId: course.id, name: groupName } },
        update: {},
        create: {
          name: groupName,
          courseId: course.id
        }
      });

      for (const [code, text] of Object.entries(options)) {
        await prisma.commentOption.upsert({
          where: { groupId_code: { groupId: group.id, code } },
          update: { text },
          create: {
            groupId: group.id,
            code,
            text
          }
        });
      }
    }
  }

  // 2. Parse Classes and Students
  const classSheets = file.SheetNames.filter(n => n !== 'Comment Banks' && n !== 'Pupil Sheets');

  // CONSTANTS FOR ANONYMIZATION
  const MALE_NAMES = ["James", "John", "Robert", "Michael", "William", "David", "Richard", "Joseph", "Thomas", "Charles", "Oliver", "George", "Harry", "Jack", "Jacob", "Noah", "Charlie", "Muhammad", "Thomas", "Oscar"];
  const FEMALE_NAMES = ["Mary", "Patricia", "Jennifer", "Linda", "Elizabeth", "Barbara", "Susan", "Jessica", "Sarah", "Karen", "Olivia", "Amelia", "Isla", "Ava", "Emily", "Isabella", "Mia", "Poppy", "Ella", "Lily"];
  const LAST_NAMES = ["Smith", "Jones", "Taylor", "Brown", "Williams", "Wilson", "Johnson", "Davies", "Robinson", "Wright", "Thompson", "Evans", "Walker", "White", "Roberts", "Green", "Hall", "Wood", "Harris", "Martin"];

  const getRandomElement = (arr: string[]) => arr[Math.floor(Math.random() * arr.length)];

  // Clear existing students to prevent duplicates
  console.log("Clearing existing student data...");
  await prisma.student.deleteMany({});

  for (const className of classSheets) {
    console.log(`Seeding Class: ${className}`);
    
    // Determine Course ID logic
    // legacy: if (classId.includes('7') || classId.includes('8') || classId.includes('9')) return classId.substring(1, 3);
    let courseName = className;
    if (className.includes('7') || className.includes('8') || className.includes('9')) {
      courseName = className.substring(1, 3);
    }
    
    // Find course
    const course = await prisma.course.findUnique({ where: { name: courseName } });
    if (!course) {
      console.warn(`Course ${courseName} not found for Class ${className}, skipping.`);
      continue;
    }

    const cls = await prisma.class.upsert({
      where: { name: className },
      update: { courseId: course.id },
      create: {
        name: className,
        courseId: course.id
      }
    });

    const sheet = file.Sheets[className];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
    // Row 0 is header.
    // [ "Family Name", "First Name", "Gender", "Form", "WP", "TH", "PS", "OA", ... ]
    
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length < 3) continue;

        // Anonymize Name
        const gender = row[2]; // "Male" / "Female"
        
        let firstName = "Student";
        if (gender === 'Male') firstName = getRandomElement(MALE_NAMES);
        else if (gender === 'Female') firstName = getRandomElement(FEMALE_NAMES);
        else firstName = getRandomElement([...MALE_NAMES, ...FEMALE_NAMES]);

        const lastName = getRandomElement(LAST_NAMES);

        // row[4] = WP code (e.g. "WP-H"). We want "H".
        // Helper to strip prefix
        // const extractCode = (val: string) => val && val.includes('-') ? val.split('-')[1] : val;

        // const wpCode = extractCode(row[4]);
        // const thCode = extractCode(row[5]);
        // const psCode = extractCode(row[6]);
        // const oaCode = extractCode(row[7]);

        // User requested to clear comment data for pupils
        const wpCode = null;
        const thCode = null;
        const psCode = null;
        const oaCode = null;

        await prisma.student.create({
            data: {
                firstName,
                lastName,
                gender,
                classId: cls.id,
                wpCode,
                thCode,
                psCode,
                oaCode
            }
        });
    }
  }
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
