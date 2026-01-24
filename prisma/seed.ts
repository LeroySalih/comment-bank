import { PrismaClient } from '@prisma/client';
import fs from 'fs';
import path from 'path';
import { hash } from 'bcryptjs';
import { encrypt } from '../lib/encryption';

const prisma = new PrismaClient();
const PUPIL_LIST_PATH = './data/pupil-list.md';

async function main() {
  // Clear existing data
  console.log("Clearing existing data...");
  await (prisma as any).pupilCode.deleteMany({});
  await (prisma as any).assignment.deleteMany({});
  await (prisma as any).pupil.deleteMany({});
  await prisma.class.deleteMany({});
  await prisma.commentOption.deleteMany({});
  await prisma.commentGroup.deleteMany({});
  await (prisma as any).subject.deleteMany({});
  await (prisma as any).user.deleteMany({ where: { username: { in: ['admin', 'leroysalih'] } } });

  // Seed Roles
  const adminRole = await prisma.role.upsert({ where: { name: 'admin' }, update: {}, create: { name: 'admin' } });
  const hodRole = await prisma.role.upsert({ where: { name: 'hod' }, update: {}, create: { name: 'hod' } });
  const teacherRole = await prisma.role.upsert({ where: { name: 'teacher' }, update: {}, create: { name: 'teacher' } });

  // Seed Admin User
  const password = await hash('password', 10);
  const adminUser = await prisma.user.upsert({
    where: { username: 'admin' },
    update: { password, roles: { connect: { id: adminRole.id } } },
    create: { username: 'admin', password, roles: { connect: { id: adminRole.id } } }
  });

  // Seed HOD User
  await prisma.user.upsert({
    where: { username: 'leroysalih' },
    update: { password, roles: { connect: { id: hodRole.id } } },
    create: { username: 'leroysalih', password, roles: { connect: { id: hodRole.id } } }
  });

  // Create a default Subject
  // Create a default Subject
  const subject = await (prisma as any).subject.create({
    data: {
      code: "7CS",
      title: "Computer Science",
      users: {
        connect: { id: adminUser.id }
      },
      commentGroups: {
        create: [
          {
            name: "WP",
            displayOrder: 0,
            options: {
              create: [
                { code: "H", text: "<Name> has shown excellent understanding of <Subject> this term." },
                { code: "M", text: "<Name> is making good progress in <Subject>." },
                { code: "L", text: "<Name> needs to focus more on <Subject> concepts." }
              ]
            }
          },
          {
            name: "TH",
            displayOrder: 1,
            options: {
              create: [
                { code: "H", text: "<Name> has demonstrated exceptional critical thinking." },
                { code: "M", text: "<Name> shows good thinking skills." },
                { code: "L", text: "<Name> is encouraged to develop more independent thinking skills." }
              ]
            }
          }
        ]
      }
    }
  });

  // Parse Pupil List
  const mdContent = fs.readFileSync(PUPIL_LIST_PATH, 'utf-8');
  const lines = mdContent.split('\n');
  
  // Skip header and separator
  const dataLines = lines.filter(l => l.includes('|') && !l.includes('Adm. Number') && !l.includes(':---'));

  console.log(`Found ${dataLines.length} pupils in ${PUPIL_LIST_PATH}`);

  for (const line of dataLines) {
    const parts = line.split('|').map(p => p.trim()).filter(p => p !== '');
    if (parts.length < 5) continue;

    const admissionNumber = parts[0];
    const lastName = parts[1];
    const firstName = parts[2];
    const gender = parts[3];
    const className = parts[4];

    // Create/Update Pupil
    const pupil = await (prisma as any).pupil.upsert({
      where: { admissionNumber },
      update: { 
        firstName: encrypt(firstName), 
        lastName: encrypt(lastName), 
        gender, 
        isActive: true 
      },
      create: { 
        admissionNumber, 
        firstName: encrypt(firstName), 
        lastName: encrypt(lastName), 
        gender, 
        isActive: true 
      }
    });

    // Create/Update Class
    const cls = await (prisma as any).class.upsert({
      where: { name: className },
      update: {},
      create: { 
        name: className, 
        subjectId: subject.id,
        year: className.substring(0, 1) 
      }
    });

    // Create Assignment
    await (prisma as any).assignment.create({
      data: {
        pupilId: pupil.admissionNumber,
        classId: cls.id,
        targetLevel: "7H",
        eoyLevel: "7M"
      }
    });
  }

  console.log('Seeding completed successfully');
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
