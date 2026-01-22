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

  // Clear existing data to prevent duplicates and ensure fresh test data
  console.log("Clearing existing data...");
  // Delete in order of dependencies
  await prisma.commentOption.deleteMany({});
  await prisma.commentGroup.deleteMany({});
  await prisma.student.deleteMany({});
  await prisma.class.deleteMany({});
  await prisma.course.deleteMany({});
  // Clean up specific test users to ensure fresh state (don't delete all users)
  await prisma.user.deleteMany({ where: { username: { in: ['admin', 'leroysalih'] } } });

  // Seed Roles
  const adminRole = await prisma.role.upsert({ where: { name: 'admin' }, update: {}, create: { name: 'admin' } });
  const hodRole = await prisma.role.upsert({ where: { name: 'hod' }, update: {}, create: { name: 'hod' } });
  const teacherRole = await prisma.role.upsert({ where: { name: 'teacher' }, update: {}, create: { name: 'teacher' } });

  // Seed Admin User
  const password = await hash('password', 10);
  const adminUser = await prisma.user.upsert({
    where: { username: 'admin' },
    update: {
        password, // Ensure password is reset to known value
        roles: { connect: { id: adminRole.id } }
    },
    create: {
      username: 'admin',
      password,
      roles: { connect: { id: adminRole.id } }
    }
  });
  console.log('Seeded User: admin / password');

  // Seed HOD User (for tests)
  await prisma.user.upsert({
    where: { username: 'leroysalih' },
    update: {
        password,
        roles: { connect: { id: hodRole.id } }
    },
    create: {
      username: 'leroysalih',
      password,
      roles: { connect: { id: hodRole.id } }
    }
  });
  console.log('Seeded User: leroysalih / password');

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

    // Inject Variable Examples for Testing
  if (!courses['11CS']) {
    courses['11CS'] = { groups: {} };
  }
  
  courses['11CS'].groups['WP'] = {
      'H': '<Name> has shown excellent understanding of <Subject> this term. <He> consistently produces high quality work.',
      'M': '<Name> is making good progress in <Subject>. <His> work is gaining consistency.',
      'L': '<Name> needs to focus more on <Subject> concepts. <He> struggles with basic terminology.'
  };
  courses['11CS'].groups['TH'] = {
      'H': '<Name> has demonstrated exceptional critical thinking. <He> can analyze complex problems with ease.',
      'M': '<Name> shows good thinking skills, though sometimes needs guidance with more abstract concepts.',
      'L': '<Name> is encouraged to develop more independent thinking skills.'
  };
  courses['11CS'].groups['PS'] = {
      'H': '<Name> is a natural problem solver. <He> approaches new challenges with confidence.',
      'M': '<Name> usually finds solutions to problems, though <his> method can be refined.',
      'L': '<Name> finds problem solving difficult and often requires step-by-step support.'
  };
  courses['11CS'].groups['OA'] = {
      'H': 'Overall, <Name> has had an outstanding term in <Subject>.',
      'M': 'Overall, it has been a positive term for <Name>.',
      'L': 'Overall, <Name> must put in more effort next term to achieve <his> target level.'
  };

  // Insert Courses and Comments
  for (const [courseName, data] of Object.entries(courses)) {
    console.log(`Seeding Course: ${courseName}`);
    
    // Simple subject mapping for test data
    let subject = "Computer Science";
    if (courseName.includes("EN")) subject = "English";
    if (courseName.includes("MA")) subject = "Maths";

    const course = await prisma.course.upsert({
      where: { name: courseName },
      update: { 
        studiedComment: data.studied, 
        ownerId: adminUser.id,
        subject 
      },
      create: { 
        name: courseName,
        studiedComment: data.studied,
        ownerId: adminUser.id,
        subject
      }
    });

    let groupIndex = 0;
    for (const [groupName, options] of Object.entries(data.groups)) {
      const group = await prisma.commentGroup.upsert({
        where: { courseId_name: { courseId: course.id, name: groupName } },
        update: { displayOrder: groupIndex },
        create: {
          name: groupName,
          courseId: course.id,
          displayOrder: groupIndex
        }
      });
      groupIndex++;

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
  
  // Helper to generate levels 1L to 9H
  const LEVELS = ['1','2','3','4','5','6','7','8','9'];
  const SUBLEVELS = ['L', 'M', 'H'];
  const getRandomLevel = () => `${getRandomElement(LEVELS)}${getRandomElement(SUBLEVELS)}`;


  for (const className of classSheets) {
    console.log(`Seeding Class: ${className}`);
    
    // Determine Course ID logic
    let courseName = className;
    if (className.includes('7') || className.includes('8') || className.includes('9')) {
      courseName = className.substring(1, 3);
    }
    
    // Extract Year from Class Name (e.g. "7CP" -> "7", "11CS" -> "11")
    const yearMatch = className.match(/^\d+/);
    const year = yearMatch ? yearMatch[0] : null;

    // Find course
    const course = await prisma.course.findUnique({ where: { name: courseName } });
    if (!course) {
      console.warn(`Course ${courseName} not found for Class ${className}, skipping.`);
      continue;
    }

    const cls = await prisma.class.upsert({
      where: { name: className },
      update: { 
        courseId: course.id,
        year 
      },
      create: {
        name: className,
        courseId: course.id,
        year
      }
    });

    const sheet = file.Sheets[className];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
    
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

        // User requested to populate fields
        // Generate random levels
        const eoyLevel = getRandomLevel();
        const targetLevel = getRandomLevel();

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
                oaCode,
                eoyLevel,
                targetLevel
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
