import 'dotenv/config';
import { PrismaClient } from '@prisma/client';
import { PrismaPg } from '@prisma/adapter-pg';
import { Pool } from 'pg';
import fs from 'fs';
import path from 'path';
import { hash } from 'bcryptjs';
import { encrypt } from '../lib/encryption';
import { createId as cuid } from '@paralleldrive/cuid2';

const pool = new Pool({ connectionString: process.env.DATABASE_URL! });
const adapter = new PrismaPg(pool);
const prisma = new PrismaClient({ adapter });
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
  await (prisma as any).user.deleteMany({ where: { username: { in: ['admin', 'leroysalih', 'teacher', 'teacher2', 'teacher3', 'teacher4'] } } });

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

  const teacher4User = await prisma.user.upsert({
    where: { username: 'teacher4' },
    update: { password, Role: { connect: { id: teacherRole.id } } },
    create: { id: cuid(), username: 'teacher4', password, Role: { connect: { id: teacherRole.id } } }
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
    // 25-7A
    { admNo: "12345", lastName: "Smith", firstName: "John", gender: "M", className: "25-7A", form: "25-7A", target: "7H", eoy: "7M" },
    { admNo: "12346", lastName: "Johnson", firstName: "Emma", gender: "F", className: "25-7A", form: "25-7A", target: "7H", eoy: "7H" },
    { admNo: "12350", lastName: "Davis", firstName: "Noah", gender: "M", className: "25-7A", form: "25-7A", target: "7M", eoy: "7M" },
    { admNo: "12351", lastName: "Miller", firstName: "Olivia", gender: "F", className: "25-7A", form: "25-7A", target: "7H", eoy: "7M" },
    { admNo: "12352", lastName: "Wilson", firstName: "James", gender: "M", className: "25-7A", form: "25-7A", target: "7M", eoy: "7L" },
    { admNo: "12353", lastName: "Moore", firstName: "Ava", gender: "F", className: "25-7A", form: "25-7A", target: "7L", eoy: "7M" },
    { admNo: "12354", lastName: "Taylor", firstName: "William", gender: "M", className: "25-7A", form: "25-7A", target: "7H", eoy: "7H" },
    { admNo: "12355", lastName: "Anderson", firstName: "Isabella", gender: "F", className: "25-7A", form: "25-7A", target: "7M", eoy: "7M" },

    // 25-7B
    { admNo: "12347", lastName: "Williams", firstName: "Oliver", gender: "M", className: "25-7B", form: "25-7B", target: "7M", eoy: "7M" },
    { admNo: "12348", lastName: "Brown", firstName: "Sophia", gender: "F", className: "25-7B", form: "25-7B", target: "7H", eoy: "7M" },
    { admNo: "12356", lastName: "Thomas", firstName: "Ethan", gender: "M", className: "25-7B", form: "25-7B", target: "7L", eoy: "7M" },
    { admNo: "12357", lastName: "Jackson", firstName: "Mia", gender: "F", className: "25-7B", form: "25-7B", target: "7M", eoy: "7H" },
    { admNo: "12358", lastName: "White", firstName: "Lucas", gender: "M", className: "25-7B", form: "25-7B", target: "7H", eoy: "7H" },
    { admNo: "12359", lastName: "Harris", firstName: "Charlotte", gender: "F", className: "25-7B", form: "25-7B", target: "7M", eoy: "7M" },
    { admNo: "12360", lastName: "Martin", firstName: "Benjamin", gender: "M", className: "25-7B", form: "25-7B", target: "7L", eoy: "7L" },
    { admNo: "12361", lastName: "Thompson", firstName: "Amelia", gender: "F", className: "25-7B", form: "25-7B", target: "7H", eoy: "7M" },

    // 25-7C
    { admNo: "12349", lastName: "Jones", firstName: "Liam", gender: "M", className: "25-7C", form: "25-7C", target: "7M", eoy: "7H" },
    { admNo: "12362", lastName: "Garcia", firstName: "Harper", gender: "F", className: "25-7C", form: "25-7C", target: "7H", eoy: "7H" },
    { admNo: "12363", lastName: "Martinez", firstName: "Alexander", gender: "M", className: "25-7C", form: "25-7C", target: "7M", eoy: "7M" },
    { admNo: "12364", lastName: "Robinson", firstName: "Evelyn", gender: "F", className: "25-7C", form: "25-7C", target: "7L", eoy: "7M" },
    { admNo: "12365", lastName: "Clark", firstName: "Henry", gender: "M", className: "25-7C", form: "25-7C", target: "7H", eoy: "7M" },
    { admNo: "12366", lastName: "Rodriguez", firstName: "Ella", gender: "F", className: "25-7C", form: "25-7C", target: "7M", eoy: "7M" },
    { admNo: "12367", lastName: "Lewis", firstName: "Sebastian", gender: "M", className: "25-7C", form: "25-7C", target: "7M", eoy: "7L" },
    { admNo: "12368", lastName: "Lee", firstName: "Scarlett", gender: "F", className: "25-7C", form: "25-7C", target: "7H", eoy: "7H" },
    { admNo: "12369", lastName: "Walker", firstName: "Jack", gender: "M", className: "25-7C", form: "25-7C", target: "7L", eoy: "7M" },

    // 25-7D
    { admNo: "12370", lastName: "Hall", firstName: "Daniel", gender: "M", className: "25-7D", form: "25-7D", target: "7H", eoy: "7H" },
    { admNo: "12371", lastName: "Allen", firstName: "Grace", gender: "F", className: "25-7D", form: "25-7D", target: "7M", eoy: "7M" },
    { admNo: "12372", lastName: "Young", firstName: "Matthew", gender: "M", className: "25-7D", form: "25-7D", target: "7H", eoy: "7M" },
    { admNo: "12373", lastName: "King", firstName: "Chloe", gender: "F", className: "25-7D", form: "25-7D", target: "7M", eoy: "7H" },
    { admNo: "12374", lastName: "Wright", firstName: "Samuel", gender: "M", className: "25-7D", form: "25-7D", target: "7L", eoy: "7M" },
    { admNo: "12375", lastName: "Scott", firstName: "Lily", gender: "F", className: "25-7D", form: "25-7D", target: "7H", eoy: "7H" },
    { admNo: "12376", lastName: "Green", firstName: "Oscar", gender: "M", className: "25-7D", form: "25-7D", target: "7M", eoy: "7L" },
    { admNo: "12377", lastName: "Adams", firstName: "Ruby", gender: "F", className: "25-7D", form: "25-7D", target: "7M", eoy: "7M" },
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
          year: pupil.className.split('-')[1]?.substring(0, 1) || '7'
        }
      });
      classMap.set(pupil.className, cls);
    }

    const cls = classMap.get(pupil.className);

    // Create Assignment with sample linked data
    const grades = ["A*", "A", "B", "C", "D", "E", "F", "NA"];
    const randomGrade = () => grades[Math.floor(Math.random() * grades.length)];

    await (prisma as any).assignment.create({
      data: {
        id: cuid(),
        pupilId: pupil.admNo,
        classId: cls.id,
        targetLevel: pupil.target,
        eoyLevel: pupil.eoy,
        linkedData: {
          behaviour: randomGrade(),
          effort: randomGrade(),
          homework: randomGrade(),
        }
      }
    });
  }

  // Assign teachers to classes
  const classAssignments: [string, any][] = [
    ["25-7A", teacherUser],
    ["25-7B", teacher2User],
    ["25-7C", teacher3User],
    ["25-7D", teacher4User],
  ];

  for (const [className, teacher] of classAssignments) {
    const cls = classMap.get(className);
    if (cls) {
      await (prisma as any).class.update({
        where: { id: cls.id },
        data: { User: { connect: { id: teacher.id } } }
      });
      console.log(`Assigned ${teacher.username} to class: ${className}`);
    }
  }

  // ============================================================================
  // Seed Common Comment Groups (CCGs)
  // ============================================================================
  console.log('\nSeeding Common Comment Groups...');

  // Clear existing CCG data
  await (prisma as any).commonPupilCode.deleteMany({});
  await (prisma as any).commonCommentOption.deleteMany({});
  await (prisma as any).commonCommentGroup.deleteMany({});
  await (prisma as any).appSetting.deleteMany({});

  const ccgData = [
    {
      name: "Academic",
      title: "Academic Performance",
      displayOrder: 0,
      isLinked: false,
      linkedField: null as string | null,
      options: [
        { code: "H", text: "<Name> has demonstrated an excellent level of academic performance in <Subject> this term.", displayOrder: 0 },
        { code: "M", text: "<Name> has demonstrated a good level of academic performance in <Subject> this term.", displayOrder: 1 },
        { code: "L", text: "<Name> has found <Subject> challenging this term and needs to improve <his> academic performance.", displayOrder: 2 },
      ]
    },
    {
      name: "Effort",
      title: "Effort",
      displayOrder: 1,
      isLinked: true,
      linkedField: "effort",
      options: [
        { code: "A*", text: "<He> consistently puts in outstanding effort in lessons, going above and beyond expectations.", displayOrder: 0 },
        { code: "A", text: "<He> consistently puts in excellent effort in lessons.", displayOrder: 1 },
        { code: "B", text: "<He> generally puts in good effort in lessons.", displayOrder: 2 },
        { code: "C", text: "<He> puts in satisfactory effort in lessons but could do more.", displayOrder: 3 },
        { code: "D", text: "<He> needs to put in more consistent effort in lessons.", displayOrder: 4 },
        { code: "E", text: "<He> needs to significantly improve <his> effort in lessons.", displayOrder: 5 },
        { code: "F", text: "<He> rarely puts in the required effort in lessons.", displayOrder: 6 },
        { code: "NA", text: "", displayOrder: 7 },
      ]
    },
    {
      name: "Behaviour",
      title: "Behaviour",
      displayOrder: 2,
      isLinked: true,
      linkedField: "behaviour",
      options: [
        { code: "A*", text: "<His> behaviour in class is exemplary and a model for others.", displayOrder: 0 },
        { code: "A", text: "<His> behaviour in class is exemplary.", displayOrder: 1 },
        { code: "B", text: "<His> behaviour in class is generally good.", displayOrder: 2 },
        { code: "C", text: "<His> behaviour in class is satisfactory.", displayOrder: 3 },
        { code: "D", text: "<He> needs to improve <his> behaviour in class.", displayOrder: 4 },
        { code: "E", text: "<He> needs to significantly improve <his> behaviour in class.", displayOrder: 5 },
        { code: "F", text: "<His> behaviour in class is a serious concern.", displayOrder: 6 },
        { code: "NA", text: "", displayOrder: 7 },
      ]
    },
    {
      name: "Homework",
      title: "Homework",
      displayOrder: 3,
      isLinked: true,
      linkedField: "homework",
      options: [
        { code: "A*", text: "<He> always completes homework to an outstanding standard and on time.", displayOrder: 0 },
        { code: "A", text: "<He> always completes homework to a high standard and on time.", displayOrder: 1 },
        { code: "B", text: "<He> usually completes homework to a good standard and on time.", displayOrder: 2 },
        { code: "C", text: "<He> usually completes homework on time.", displayOrder: 3 },
        { code: "D", text: "<He> needs to ensure that homework is completed on time.", displayOrder: 4 },
        { code: "E", text: "<He> frequently fails to complete homework on time or to a satisfactory standard.", displayOrder: 5 },
        { code: "F", text: "<He> rarely completes homework.", displayOrder: 6 },
        { code: "NA", text: "", displayOrder: 7 },
      ]
    },
    {
      name: "Overall",
      title: "Overall",
      displayOrder: 4,
      isLinked: false,
      linkedField: null as string | null,
      options: [
        { code: "H", text: "Overall, <Name> is making excellent progress and should continue to work at this level.", displayOrder: 0 },
        { code: "M", text: "Overall, <Name> is making good progress and should continue to build on this.", displayOrder: 1 },
        { code: "L", text: "Overall, <Name> needs to focus on improving <his> effort and engagement to make better progress.", displayOrder: 2 },
      ]
    },
  ];

  for (const group of ccgData) {
    await (prisma as any).commonCommentGroup.create({
      data: {
        id: cuid(),
        name: group.name,
        title: group.title,
        displayOrder: group.displayOrder,
        isLinked: group.isLinked,
        linkedField: group.linkedField,
        CommonCommentOption: {
          create: group.options.map(opt => ({
            id: cuid(),
            code: opt.code,
            text: opt.text,
            displayOrder: opt.displayOrder,
          }))
        }
      }
    });
  }

  // Seed comment format template
  await (prisma as any).appSetting.upsert({
    where: { key: 'comment_format_template' },
    update: { value: '<Academic>\n\n<Effort> <Behaviour> <Homework>\n\n<SCG>\n\n<Overall>' },
    create: { key: 'comment_format_template', value: '<Academic>\n\n<Effort> <Behaviour> <Homework>\n\n<SCG>\n\n<Overall>' }
  });

  // Set sample commentFormat on subjects
  await (prisma as any).subject.update({
    where: { id: subject7CS.id },
    data: { commentFormat: 'WP TH' }
  });
  await (prisma as any).subject.update({
    where: { id: subject7DT.id },
    data: { commentFormat: 'PS DS' }
  });
  await (prisma as any).subject.update({
    where: { id: subject8CS.id },
    data: { commentFormat: 'PR TH' }
  });

  console.log('Common Comment Groups seeded successfully');

  console.log('\n=== Seeding Summary ===');
  console.log(`Created ${pupilData.length} pupils`);
  console.log(`Created 4 classes: 25-7A, 25-7B, 25-7C, 25-7D`);
  console.log(`Created 3 subjects: 7CS, 7DT, 8CS`);
  console.log(`Assigned all subjects to leroysalih (HOD)`);
  console.log(`Created 6 users: admin, leroysalih, teacher, teacher2, teacher3, teacher4`);
  console.log(`Created 5 Common Comment Groups with options`);
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
