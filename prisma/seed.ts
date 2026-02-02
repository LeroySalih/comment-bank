import { PrismaClient } from '@prisma/client';
import fs from 'fs';
import path from 'path';
import { hash } from 'bcryptjs';
import { encrypt } from '../lib/encryption';
import { createId as cuid } from '@paralleldrive/cuid2';

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
  await (prisma as any).user.deleteMany({ where: { username: { in: ['admin', 'leroysalih', 'teacher', 'teacher2', 'teacher3'] } } });

  // Seed Roles
  const adminRole = await prisma.role.upsert({ where: { name: 'admin' }, update: {}, create: { id: cuid(), name: 'admin' } });
  const hodRole = await prisma.role.upsert({ where: { name: 'hod' }, update: {}, create: { id: cuid(), name: 'hod' } });
  const teacherRole = await prisma.role.upsert({ where: { name: 'teacher' }, update: {}, create: { id: cuid(), name: 'teacher' } });

  // Seed Admin User
  const password = await hash('password', 10);
  const adminUser = await prisma.user.upsert({
    where: { username: 'admin' },
    update: { password, Role: { connect: { id: adminRole.id } } },
    create: { id: cuid(), username: 'admin', password, Role: { connect: { id: adminRole.id } } }
  });

  // Seed HOD User (leroysalih)
  const hodUser = await prisma.user.upsert({
    where: { username: 'leroysalih' },
    update: { password, Role: { connect: { id: hodRole.id } } },
    create: { id: cuid(), username: 'leroysalih', password, Role: { connect: { id: hodRole.id } } }
  });

  // Seed Teacher Users
  const teacherUser = await prisma.user.upsert({
    where: { username: 'teacher' },
    update: { password, Role: { connect: { id: teacherRole.id } } },
    create: { id: cuid(), username: 'teacher', password, Role: { connect: { id: teacherRole.id } } }
  });

  const teacher2User = await prisma.user.upsert({
    where: { username: 'teacher2' },
    update: { password, Role: { connect: { id: teacherRole.id } } },
    create: { id: cuid(), username: 'teacher2', password, Role: { connect: { id: teacherRole.id } } }
  });

  const teacher3User = await prisma.user.upsert({
    where: { username: 'teacher3' },
    update: { password, Role: { connect: { id: teacherRole.id } } },
    create: { id: cuid(), username: 'teacher3', password, Role: { connect: { id: teacherRole.id } } }
  });

  // Create Subjects with Comment Groups
  // Subject 1: 7CS - Computer Science (PRESERVE EXISTING COMMENT GROUPS)
  const subject7CS = await (prisma as any).subject.create({
    data: {
      id: cuid(),
      code: "7CS",
      title: "Computer Science",
      User: {
        connect: { id: hodUser.id } // Assign to leroysalih (HOD)
      },
      CommentGroup: {
        create: [
          {
            id: cuid(),
            name: "WP",
            title: "Working Progress",
            displayOrder: 0,
            CommentOption: {
              create: [
                { id: cuid(), code: "H", text: "<Name> has shown excellent understanding of <Subject> this term.", displayOrder: 0 },
                { id: cuid(), code: "M", text: "<Name> is making good progress in <Subject>.", displayOrder: 1 },
                { id: cuid(), code: "L", text: "<Name> needs to focus more on <Subject> concepts.", displayOrder: 2 }
              ]
            }
          },
          {
            id: cuid(),
            name: "TH",
            title: "Thinking",
            displayOrder: 1,
            CommentOption: {
              create: [
                { id: cuid(), code: "H", text: "<Name> has demonstrated exceptional critical thinking.", displayOrder: 0 },
                { id: cuid(), code: "M", text: "<Name> shows good thinking skills.", displayOrder: 1 },
                { id: cuid(), code: "L", text: "<Name> is encouraged to develop more independent thinking skills.", displayOrder: 2 }
              ]
            }
          }
        ]
      }
    }
  });

  // Subject 2: 7DT - Design & Technology
  const subject7DT = await (prisma as any).subject.create({
    data: {
      id: cuid(),
      code: "7DT",
      title: "Design & Technology",
      User: {
        connect: { id: hodUser.id } // Assign to leroysalih (HOD)
      },
      CommentGroup: {
        create: [
          {
            id: cuid(),
            name: "PS",
            title: "Practical Skills",
            displayOrder: 0,
            CommentOption: {
              create: [
                { id: cuid(), code: "H", text: "<Name> demonstrates excellent practical skills in <Subject>.", displayOrder: 0 },
                { id: cuid(), code: "M", text: "<Name> shows good practical ability in <Subject>.", displayOrder: 1 },
                { id: cuid(), code: "L", text: "<Name> needs to develop <his/her> practical skills further.", displayOrder: 2 }
              ]
            }
          },
          {
            id: cuid(),
            name: "DS",
            title: "Design Skills",
            displayOrder: 1,
            CommentOption: {
              create: [
                { id: cuid(), code: "H", text: "<Name> produces creative and innovative designs.", displayOrder: 0 },
                { id: cuid(), code: "M", text: "<Name> creates appropriate designs for the task.", displayOrder: 1 },
                { id: cuid(), code: "L", text: "<Name> should focus on developing more creative design solutions.", displayOrder: 2 }
              ]
            }
          }
        ]
      }
    }
  });

  // Subject 3: 8CS - Computer Science Year 8
  const subject8CS = await (prisma as any).subject.create({
    data: {
      id: cuid(),
      code: "8CS",
      title: "Computer Science Year 8",
      User: {
        connect: { id: hodUser.id } // Assign to leroysalih (HOD)
      },
      CommentGroup: {
        create: [
          {
            id: cuid(),
            name: "PR",
            title: "Programming",
            displayOrder: 0,
            CommentOption: {
              create: [
                { id: cuid(), code: "H", text: "<Name> excels at programming and problem-solving.", displayOrder: 0 },
                { id: cuid(), code: "M", text: "<Name> is developing good programming skills.", displayOrder: 1 },
                { id: cuid(), code: "L", text: "<Name> needs more practice with programming concepts.", displayOrder: 2 }
              ]
            }
          },
          {
            id: cuid(),
            name: "TH",
            title: "Theory",
            displayOrder: 1,
            CommentOption: {
              create: [
                { id: cuid(), code: "H", text: "<Name> has excellent understanding of computing theory.", displayOrder: 0 },
                { id: cuid(), code: "M", text: "<Name> shows good grasp of theoretical concepts.", displayOrder: 1 },
                { id: cuid(), code: "L", text: "<Name> should review theoretical concepts more thoroughly.", displayOrder: 2 }
              ]
            }
          }
        ]
      }
    }
  });

  // Generate realistic pupils
  const pupilData = [
    // Year 7 Class A
    { admNo: "12345", lastName: "Smith", firstName: "John", gender: "M", className: "7A", form: "7A", target: "7H", eoy: "7M" },
    { admNo: "12346", lastName: "Johnson", firstName: "Emma", gender: "F", className: "7A", form: "7A", target: "7H", eoy: "7H" },
    { admNo: "12350", lastName: "Davis", firstName: "Noah", gender: "M", className: "7A", form: "7A", target: "7M", eoy: "7M" },
    { admNo: "12351", lastName: "Miller", firstName: "Olivia", gender: "F", className: "7A", form: "7A", target: "7H", eoy: "7M" },
    { admNo: "12352", lastName: "Wilson", firstName: "James", gender: "M", className: "7A", form: "7A", target: "7M", eoy: "7L" },
    { admNo: "12353", lastName: "Moore", firstName: "Ava", gender: "F", className: "7A", form: "7A", target: "7L", eoy: "7M" },
    { admNo: "12354", lastName: "Taylor", firstName: "William", gender: "M", className: "7A", form: "7A", target: "7H", eoy: "7H" },
    { admNo: "12355", lastName: "Anderson", firstName: "Isabella", gender: "F", className: "7A", form: "7A", target: "7M", eoy: "7M" },
    
    // Year 7 Class B
    { admNo: "12347", lastName: "Williams", firstName: "Oliver", gender: "M", className: "7B", form: "7B", target: "7M", eoy: "7M" },
    { admNo: "12348", lastName: "Brown", firstName: "Sophia", gender: "F", className: "7B", form: "7B", target: "7H", eoy: "7M" },
    { admNo: "12356", lastName: "Thomas", firstName: "Ethan", gender: "M", className: "7B", form: "7B", target: "7L", eoy: "7M" },
    { admNo: "12357", lastName: "Jackson", firstName: "Mia", gender: "F", className: "7B", form: "7B", target: "7M", eoy: "7H" },
    { admNo: "12358", lastName: "White", firstName: "Lucas", gender: "M", className: "7B", form: "7B", target: "7H", eoy: "7H" },
    { admNo: "12359", lastName: "Harris", firstName: "Charlotte", gender: "F", className: "7B", form: "7B", target: "7M", eoy: "7M" },
    { admNo: "12360", lastName: "Martin", firstName: "Benjamin", gender: "M", className: "7B", form: "7B", target: "7L", eoy: "7L" },
    { admNo: "12361", lastName: "Thompson", firstName: "Amelia", gender: "F", className: "7B", form: "7B", target: "7H", eoy: "7M" },
    
    // Year 7 Class C
    { admNo: "12349", lastName: "Jones", firstName: "Liam", gender: "M", className: "7C", form: "7C", target: "7M", eoy: "7H" },
    { admNo: "12362", lastName: "Garcia", firstName: "Harper", gender: "F", className: "7C", form: "7C", target: "7H", eoy: "7H" },
    { admNo: "12363", lastName: "Martinez", firstName: "Alexander", gender: "M", className: "7C", form: "7C", target: "7M", eoy: "7M" },
    { admNo: "12364", lastName: "Robinson", firstName: "Evelyn", gender: "F", className: "7C", form: "7C", target: "7L", eoy: "7M" },
    { admNo: "12365", lastName: "Clark", firstName: "Henry", gender: "M", className: "7C", form: "7C", target: "7H", eoy: "7M" },
    { admNo: "12366", lastName: "Rodriguez", firstName: "Ella", gender: "F", className: "7C", form: "7C", target: "7M", eoy: "7M" },
    { admNo: "12367", lastName: "Lewis", firstName: "Sebastian", gender: "M", className: "7C", form: "7C", target: "7M", eoy: "7L" },
    { admNo: "12368", lastName: "Lee", firstName: "Scarlett", gender: "F", className: "7C", form: "7C", target: "7H", eoy: "7H" },
    { admNo: "12369", lastName: "Walker", firstName: "Jack", gender: "M", className: "7C", form: "7C", target: "7L", eoy: "7M" },
  ];

  const classMap = new Map();

  // Create pupils and classes
  for (const pupil of pupilData) {
    // Create/Update Pupil
    await (prisma as any).pupil.upsert({
      where: { admissionNumber: pupil.admNo },
      update: { 
        firstName: encrypt(pupil.firstName), 
        lastName: encrypt(pupil.lastName), 
        gender: pupil.gender,
        form: pupil.form,
        isActive: true 
      },
      create: { 
        admissionNumber: pupil.admNo, 
        firstName: encrypt(pupil.firstName), 
        lastName: encrypt(pupil.lastName), 
        gender: pupil.gender,
        form: pupil.form,
        isActive: true 
      }
    });

    // Create/Update Class for 7CS
    if (!classMap.has(pupil.className)) {
      const cls = await (prisma as any).class.upsert({
        where: { name: pupil.className },
        update: {},
        create: { 
          id: cuid(),
          name: pupil.className, 
          subjectId: subject7CS.id,
          year: pupil.className.substring(0, 1) 
        }
      });
      classMap.set(pupil.className, cls);
    }

    const cls = classMap.get(pupil.className);

    // Create Assignment
    await (prisma as any).assignment.create({
      data: {
        id: cuid(),
        pupilId: pupil.admNo,
        classId: cls.id,
        targetLevel: pupil.target,
        eoyLevel: pupil.eoy
      }
    });
  }

  // Assign teachers to classes
  const class7A = classMap.get("7A");
  const class7B = classMap.get("7B");
  const class7C = classMap.get("7C");

  if (class7A) {
    await (prisma as any).class.update({
      where: { id: class7A.id },
      data: {
        User: {
          connect: { id: teacherUser.id }
        }
      }
    });
    console.log(`Assigned teacher to class: 7A`);
  }

  if (class7B) {
    await (prisma as any).class.update({
      where: { id: class7B.id },
      data: {
        User: {
          connect: { id: teacher2User.id }
        }
      }
    });
    console.log(`Assigned teacher2 to class: 7B`);
  }

  if (class7C) {
    await (prisma as any).class.update({
      where: { id: class7C.id },
      data: {
        User: {
          connect: { id: teacher3User.id }
        }
      }
    });
    console.log(`Assigned teacher3 to class: 7C`);
  }

  console.log('\n=== Seeding Summary ===');
  console.log(`Created ${pupilData.length} pupils`);
  console.log(`Created 3 classes: 7A, 7B, 7C`);
  console.log(`Created 3 subjects: 7CS, 7DT, 8CS`);
  console.log(`Assigned all subjects to leroysalih (HOD)`);
  console.log(`Created 5 users: admin, leroysalih, teacher, teacher2, teacher3`);
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
