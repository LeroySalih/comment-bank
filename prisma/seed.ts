import 'dotenv/config';
import { Pool } from 'pg';
import { hash } from 'bcryptjs';
import { encrypt } from '../lib/encryption';
import { createId } from '@paralleldrive/cuid2';

const pool = new Pool({ connectionString: process.env.DATABASE_URL! });

async function main() {
  const client = await pool.connect();

  try {
    // Clear existing data
    console.log("Clearing existing data...");
    await client.query(`DELETE FROM "PupilCode"`);
    await client.query(`DELETE FROM "Assignment"`);
    await client.query(`DELETE FROM "Pupil"`);
    await client.query(`DELETE FROM "_ClassToUser"`);
    await client.query(`DELETE FROM "_PupilToClass"`);
    await client.query(`DELETE FROM "Class"`);
    await client.query(`DELETE FROM "CommentOption"`);
    await client.query(`DELETE FROM "CommentGroup"`);
    await client.query(`DELETE FROM "_SubjectToUser"`);
    await client.query(`DELETE FROM "Subject"`);
    await client.query(
      `DELETE FROM "_RoleToUser" WHERE "B" IN (
        SELECT id FROM "User" WHERE username IN ('admin','leroysalih','teacher','teacher2','teacher3','teacher4')
      )`
    );
    await client.query(
      `DELETE FROM "User" WHERE username IN ('admin','leroysalih','teacher','teacher2','teacher3','teacher4')`
    );

    // Seed Roles
    const adminRoleId = createId();
    const hodRoleId = createId();
    const teacherRoleId = createId();

    await client.query(
      `INSERT INTO "Role" (id, name) VALUES ($1, $2) ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name RETURNING id`,
      [adminRoleId, 'admin']
    );
    const { rows: adminRoleRows } = await client.query(
      `SELECT id FROM "Role" WHERE name = 'admin'`
    );
    const actualAdminRoleId: string = adminRoleRows[0].id;

    await client.query(
      `INSERT INTO "Role" (id, name) VALUES ($1, $2) ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name RETURNING id`,
      [hodRoleId, 'hod']
    );
    const { rows: hodRoleRows } = await client.query(
      `SELECT id FROM "Role" WHERE name = 'hod'`
    );
    const actualHodRoleId: string = hodRoleRows[0].id;

    await client.query(
      `INSERT INTO "Role" (id, name) VALUES ($1, $2) ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name RETURNING id`,
      [teacherRoleId, 'teacher']
    );
    const { rows: teacherRoleRows } = await client.query(
      `SELECT id FROM "Role" WHERE name = 'teacher'`
    );
    const actualTeacherRoleId: string = teacherRoleRows[0].id;

    // Seed Users
    const password = await hash('password', 10);

    const adminUserId = createId();
    await client.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)
       ON CONFLICT (username) DO UPDATE SET password = EXCLUDED.password, "isActive" = true`,
      [adminUserId, 'admin', password]
    );
    const { rows: adminRows } = await client.query(
      `SELECT id FROM "User" WHERE username = 'admin'`
    );
    const adminUser = adminRows[0];
    await client.query(
      `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [actualAdminRoleId, adminUser.id]
    );

    const hodUserId = createId();
    await client.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)
       ON CONFLICT (username) DO UPDATE SET password = EXCLUDED.password, "isActive" = true`,
      [hodUserId, 'leroysalih', password]
    );
    const { rows: hodRows } = await client.query(
      `SELECT id FROM "User" WHERE username = 'leroysalih'`
    );
    const hodUser = hodRows[0];
    await client.query(
      `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [actualHodRoleId, hodUser.id]
    );

    const teacherUserId = createId();
    await client.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)
       ON CONFLICT (username) DO UPDATE SET password = EXCLUDED.password, "isActive" = true`,
      [teacherUserId, 'teacher', password]
    );
    const { rows: teacherRows } = await client.query(
      `SELECT id FROM "User" WHERE username = 'teacher'`
    );
    const teacherUser = teacherRows[0];
    await client.query(
      `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [actualTeacherRoleId, teacherUser.id]
    );

    const teacher2UserId = createId();
    await client.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)
       ON CONFLICT (username) DO UPDATE SET password = EXCLUDED.password, "isActive" = true`,
      [teacher2UserId, 'teacher2', password]
    );
    const { rows: teacher2Rows } = await client.query(
      `SELECT id FROM "User" WHERE username = 'teacher2'`
    );
    const teacher2User = teacher2Rows[0];
    await client.query(
      `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [actualTeacherRoleId, teacher2User.id]
    );

    const teacher3UserId = createId();
    await client.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)
       ON CONFLICT (username) DO UPDATE SET password = EXCLUDED.password, "isActive" = true`,
      [teacher3UserId, 'teacher3', password]
    );
    const { rows: teacher3Rows } = await client.query(
      `SELECT id FROM "User" WHERE username = 'teacher3'`
    );
    const teacher3User = teacher3Rows[0];
    await client.query(
      `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [actualTeacherRoleId, teacher3User.id]
    );

    const teacher4UserId = createId();
    await client.query(
      `INSERT INTO "User" (id, username, password, "isActive") VALUES ($1, $2, $3, true)
       ON CONFLICT (username) DO UPDATE SET password = EXCLUDED.password, "isActive" = true`,
      [teacher4UserId, 'teacher4', password]
    );
    const { rows: teacher4Rows } = await client.query(
      `SELECT id FROM "User" WHERE username = 'teacher4'`
    );
    const teacher4User = teacher4Rows[0];
    await client.query(
      `INSERT INTO "_RoleToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [actualTeacherRoleId, teacher4User.id]
    );

    // Create Subjects with Comment Groups

    // Subject 1: 7CS - Computer Science
    const subject7CSId = createId();
    await client.query(
      `INSERT INTO "Subject" (id, code, title) VALUES ($1, $2, $3)`,
      [subject7CSId, '7CS', 'Computer Science']
    );
    await client.query(
      `INSERT INTO "_SubjectToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [subject7CSId, hodUser.id]
    );

    const wpGroupId = createId();
    await client.query(
      `INSERT INTO "CommentGroup" (id, name, title, "displayOrder", "subjectId") VALUES ($1, $2, $3, $4, $5)`,
      [wpGroupId, 'WP', 'Working Progress', 0, subject7CSId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'H', '<Name> has shown excellent understanding of <Subject> this term.', 0, wpGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'M', '<Name> is making good progress in <Subject>.', 1, wpGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'L', '<Name> needs to focus more on <Subject> concepts.', 2, wpGroupId]
    );

    const thGroupId7CS = createId();
    await client.query(
      `INSERT INTO "CommentGroup" (id, name, title, "displayOrder", "subjectId") VALUES ($1, $2, $3, $4, $5)`,
      [thGroupId7CS, 'TH', 'Thinking', 1, subject7CSId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'H', '<Name> has demonstrated exceptional critical thinking.', 0, thGroupId7CS]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'M', '<Name> shows good thinking skills.', 1, thGroupId7CS]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'L', '<Name> is encouraged to develop more independent thinking skills.', 2, thGroupId7CS]
    );

    // Subject 2: 7DT - Design & Technology
    const subject7DTId = createId();
    await client.query(
      `INSERT INTO "Subject" (id, code, title) VALUES ($1, $2, $3)`,
      [subject7DTId, '7DT', 'Design & Technology']
    );
    await client.query(
      `INSERT INTO "_SubjectToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [subject7DTId, hodUser.id]
    );

    const psGroupId = createId();
    await client.query(
      `INSERT INTO "CommentGroup" (id, name, title, "displayOrder", "subjectId") VALUES ($1, $2, $3, $4, $5)`,
      [psGroupId, 'PS', 'Practical Skills', 0, subject7DTId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'H', '<Name> demonstrates excellent practical skills in <Subject>.', 0, psGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'M', '<Name> shows good practical ability in <Subject>.', 1, psGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'L', '<Name> needs to develop <his/her> practical skills further.', 2, psGroupId]
    );

    const dsGroupId = createId();
    await client.query(
      `INSERT INTO "CommentGroup" (id, name, title, "displayOrder", "subjectId") VALUES ($1, $2, $3, $4, $5)`,
      [dsGroupId, 'DS', 'Design Skills', 1, subject7DTId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'H', '<Name> produces creative and innovative designs.', 0, dsGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'M', '<Name> creates appropriate designs for the task.', 1, dsGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'L', '<Name> should focus on developing more creative design solutions.', 2, dsGroupId]
    );

    // Subject 3: 8CS - Computer Science Year 8
    const subject8CSId = createId();
    await client.query(
      `INSERT INTO "Subject" (id, code, title) VALUES ($1, $2, $3)`,
      [subject8CSId, '8CS', 'Computer Science Year 8']
    );
    await client.query(
      `INSERT INTO "_SubjectToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
      [subject8CSId, hodUser.id]
    );

    const prGroupId = createId();
    await client.query(
      `INSERT INTO "CommentGroup" (id, name, title, "displayOrder", "subjectId") VALUES ($1, $2, $3, $4, $5)`,
      [prGroupId, 'PR', 'Programming', 0, subject8CSId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'H', '<Name> excels at programming and problem-solving.', 0, prGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'M', '<Name> is developing good programming skills.', 1, prGroupId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'L', '<Name> needs more practice with programming concepts.', 2, prGroupId]
    );

    const thGroupId8CS = createId();
    await client.query(
      `INSERT INTO "CommentGroup" (id, name, title, "displayOrder", "subjectId") VALUES ($1, $2, $3, $4, $5)`,
      [thGroupId8CS, 'TH', 'Theory', 1, subject8CSId]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'H', '<Name> has excellent understanding of computing theory.', 0, thGroupId8CS]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'M', '<Name> shows good grasp of theoretical concepts.', 1, thGroupId8CS]
    );
    await client.query(
      `INSERT INTO "CommentOption" (id, code, text, "displayOrder", "groupId") VALUES ($1,$2,$3,$4,$5)`,
      [createId(), 'L', '<Name> should review theoretical concepts more thoroughly.', 2, thGroupId8CS]
    );

    // Set commentFormat on subjects
    await client.query(`UPDATE "Subject" SET "commentFormat" = $1 WHERE id = $2`, ['WP TH', subject7CSId]);
    await client.query(`UPDATE "Subject" SET "commentFormat" = $1 WHERE id = $2`, ['PS DS', subject7DTId]);
    await client.query(`UPDATE "Subject" SET "commentFormat" = $1 WHERE id = $2`, ['PR TH', subject8CSId]);

    // Generate pupils and classes
    const pupilData = [
      // 25-7A
      { admNo: "12345", lastName: "Smith",    firstName: "John",      gender: "M", className: "25-7A", form: "25-7A", target: "7H", eoy: "7M" },
      { admNo: "12346", lastName: "Johnson",  firstName: "Emma",      gender: "F", className: "25-7A", form: "25-7A", target: "7H", eoy: "7H" },
      { admNo: "12350", lastName: "Davis",    firstName: "Noah",      gender: "M", className: "25-7A", form: "25-7A", target: "7M", eoy: "7M" },
      { admNo: "12351", lastName: "Miller",   firstName: "Olivia",    gender: "F", className: "25-7A", form: "25-7A", target: "7H", eoy: "7M" },
      { admNo: "12352", lastName: "Wilson",   firstName: "James",     gender: "M", className: "25-7A", form: "25-7A", target: "7M", eoy: "7L" },
      { admNo: "12353", lastName: "Moore",    firstName: "Ava",       gender: "F", className: "25-7A", form: "25-7A", target: "7L", eoy: "7M" },
      { admNo: "12354", lastName: "Taylor",   firstName: "William",   gender: "M", className: "25-7A", form: "25-7A", target: "7H", eoy: "7H" },
      { admNo: "12355", lastName: "Anderson", firstName: "Isabella",  gender: "F", className: "25-7A", form: "25-7A", target: "7M", eoy: "7M" },
      // 25-7B
      { admNo: "12347", lastName: "Williams", firstName: "Oliver",    gender: "M", className: "25-7B", form: "25-7B", target: "7M", eoy: "7M" },
      { admNo: "12348", lastName: "Brown",    firstName: "Sophia",    gender: "F", className: "25-7B", form: "25-7B", target: "7H", eoy: "7M" },
      { admNo: "12356", lastName: "Thomas",   firstName: "Ethan",     gender: "M", className: "25-7B", form: "25-7B", target: "7L", eoy: "7M" },
      { admNo: "12357", lastName: "Jackson",  firstName: "Mia",       gender: "F", className: "25-7B", form: "25-7B", target: "7M", eoy: "7H" },
      { admNo: "12358", lastName: "White",    firstName: "Lucas",     gender: "M", className: "25-7B", form: "25-7B", target: "7H", eoy: "7H" },
      { admNo: "12359", lastName: "Harris",   firstName: "Charlotte", gender: "F", className: "25-7B", form: "25-7B", target: "7M", eoy: "7M" },
      { admNo: "12360", lastName: "Martin",   firstName: "Benjamin",  gender: "M", className: "25-7B", form: "25-7B", target: "7L", eoy: "7L" },
      { admNo: "12361", lastName: "Thompson", firstName: "Amelia",    gender: "F", className: "25-7B", form: "25-7B", target: "7H", eoy: "7M" },
      // 25-7C
      { admNo: "12349", lastName: "Jones",    firstName: "Liam",      gender: "M", className: "25-7C", form: "25-7C", target: "7M", eoy: "7H" },
      { admNo: "12362", lastName: "Garcia",   firstName: "Harper",    gender: "F", className: "25-7C", form: "25-7C", target: "7H", eoy: "7H" },
      { admNo: "12363", lastName: "Martinez", firstName: "Alexander", gender: "M", className: "25-7C", form: "25-7C", target: "7M", eoy: "7M" },
      { admNo: "12364", lastName: "Robinson", firstName: "Evelyn",    gender: "F", className: "25-7C", form: "25-7C", target: "7L", eoy: "7M" },
      { admNo: "12365", lastName: "Clark",    firstName: "Henry",     gender: "M", className: "25-7C", form: "25-7C", target: "7H", eoy: "7M" },
      { admNo: "12366", lastName: "Rodriguez",firstName: "Ella",      gender: "F", className: "25-7C", form: "25-7C", target: "7M", eoy: "7M" },
      { admNo: "12367", lastName: "Lewis",    firstName: "Sebastian", gender: "M", className: "25-7C", form: "25-7C", target: "7M", eoy: "7L" },
      { admNo: "12368", lastName: "Lee",      firstName: "Scarlett",  gender: "F", className: "25-7C", form: "25-7C", target: "7H", eoy: "7H" },
      { admNo: "12369", lastName: "Walker",   firstName: "Jack",      gender: "M", className: "25-7C", form: "25-7C", target: "7L", eoy: "7M" },
      // 25-7D
      { admNo: "12370", lastName: "Hall",     firstName: "Daniel",    gender: "M", className: "25-7D", form: "25-7D", target: "7H", eoy: "7H" },
      { admNo: "12371", lastName: "Allen",    firstName: "Grace",     gender: "F", className: "25-7D", form: "25-7D", target: "7M", eoy: "7M" },
      { admNo: "12372", lastName: "Young",    firstName: "Matthew",   gender: "M", className: "25-7D", form: "25-7D", target: "7H", eoy: "7M" },
      { admNo: "12373", lastName: "King",     firstName: "Chloe",     gender: "F", className: "25-7D", form: "25-7D", target: "7M", eoy: "7H" },
      { admNo: "12374", lastName: "Wright",   firstName: "Samuel",    gender: "M", className: "25-7D", form: "25-7D", target: "7L", eoy: "7M" },
      { admNo: "12375", lastName: "Scott",    firstName: "Lily",      gender: "F", className: "25-7D", form: "25-7D", target: "7H", eoy: "7H" },
      { admNo: "12376", lastName: "Green",    firstName: "Oscar",     gender: "M", className: "25-7D", form: "25-7D", target: "7M", eoy: "7L" },
      { admNo: "12377", lastName: "Adams",    firstName: "Ruby",      gender: "F", className: "25-7D", form: "25-7D", target: "7M", eoy: "7M" },
    ];

    const classMap = new Map<string, string>(); // className -> classId

    const grades = ["A*", "A", "B", "C", "D", "E", "F", "NA"];
    const randomGrade = () => grades[Math.floor(Math.random() * grades.length)];

    for (const pupil of pupilData) {
      // Upsert Pupil
      await client.query(
        `INSERT INTO "Pupil" ("admissionNumber", "firstName", "lastName", gender, form, "isActive")
         VALUES ($1, $2, $3, $4, $5, true)
         ON CONFLICT ("admissionNumber") DO UPDATE
           SET "firstName" = EXCLUDED."firstName",
               "lastName"  = EXCLUDED."lastName",
               gender      = EXCLUDED.gender,
               form        = EXCLUDED.form,
               "isActive"  = true`,
        [pupil.admNo, encrypt(pupil.firstName), encrypt(pupil.lastName), pupil.gender, pupil.form]
      );

      // Upsert Class (for 7CS subject)
      if (!classMap.has(pupil.className)) {
        const classId = createId();
        const year = pupil.className.split('-')[1]?.substring(0, 1) || '7';
        await client.query(
          `INSERT INTO "Class" (id, name, "subjectId", year)
           VALUES ($1, $2, $3, $4)
           ON CONFLICT (name) DO UPDATE SET name = EXCLUDED.name`,
          [classId, pupil.className, subject7CSId, year]
        );
        const { rows: clsRows } = await client.query(
          `SELECT id FROM "Class" WHERE name = $1`,
          [pupil.className]
        );
        classMap.set(pupil.className, clsRows[0].id);
      }

      const classId = classMap.get(pupil.className)!;

      // Create Assignment (this links the pupil to the class)
      await client.query(
        `INSERT INTO "Assignment" (id, "pupilId", "classId", "targetLevel", "eoyLevel", "linkedData")
         VALUES ($1, $2, $3, $4, $5, $6)
         ON CONFLICT DO NOTHING`,
        [
          createId(),
          pupil.admNo,
          classId,
          pupil.target,
          pupil.eoy,
          JSON.stringify({ behaviour: randomGrade(), effort: randomGrade(), homework: randomGrade() })
        ]
      );
    }

    // Assign teachers to classes
    const classAssignments: [string, { id: string; username: string }][] = [
      ["25-7A", teacherUser],
      ["25-7B", teacher2User],
      ["25-7C", teacher3User],
      ["25-7D", teacher4User],
    ];

    for (const [className, teacher] of classAssignments) {
      const classId = classMap.get(className);
      if (classId) {
        await client.query(
          `INSERT INTO "_ClassToUser" ("A", "B") VALUES ($1, $2) ON CONFLICT DO NOTHING`,
          [classId, teacher.id]
        );
        console.log(`Assigned ${teacher.username} to class: ${className}`);
      }
    }

    // Seed Common Comment Groups (CCGs)
    console.log('\nSeeding Common Comment Groups...');

    await client.query(`DELETE FROM "CommonPupilCode"`);
    await client.query(`DELETE FROM "CommonCommentOption"`);
    await client.query(`DELETE FROM "CommonCommentGroup"`);
    await client.query(`DELETE FROM "AppSetting"`);

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
          { code: "A",  text: "<He> consistently puts in excellent effort in lessons.", displayOrder: 1 },
          { code: "B",  text: "<He> generally puts in good effort in lessons.", displayOrder: 2 },
          { code: "C",  text: "<He> puts in satisfactory effort in lessons but could do more.", displayOrder: 3 },
          { code: "D",  text: "<He> needs to put in more consistent effort in lessons.", displayOrder: 4 },
          { code: "E",  text: "<He> needs to significantly improve <his> effort in lessons.", displayOrder: 5 },
          { code: "F",  text: "<He> rarely puts in the required effort in lessons.", displayOrder: 6 },
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
          { code: "A",  text: "<His> behaviour in class is exemplary.", displayOrder: 1 },
          { code: "B",  text: "<His> behaviour in class is generally good.", displayOrder: 2 },
          { code: "C",  text: "<His> behaviour in class is satisfactory.", displayOrder: 3 },
          { code: "D",  text: "<He> needs to improve <his> behaviour in class.", displayOrder: 4 },
          { code: "E",  text: "<He> needs to significantly improve <his> behaviour in class.", displayOrder: 5 },
          { code: "F",  text: "<His> behaviour in class is a serious concern.", displayOrder: 6 },
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
          { code: "A",  text: "<He> always completes homework to a high standard and on time.", displayOrder: 1 },
          { code: "B",  text: "<He> usually completes homework to a good standard and on time.", displayOrder: 2 },
          { code: "C",  text: "<He> usually completes homework on time.", displayOrder: 3 },
          { code: "D",  text: "<He> needs to ensure that homework is completed on time.", displayOrder: 4 },
          { code: "E",  text: "<He> frequently fails to complete homework on time or to a satisfactory standard.", displayOrder: 5 },
          { code: "F",  text: "<He> rarely completes homework.", displayOrder: 6 },
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
      const groupId = createId();
      await client.query(
        `INSERT INTO "CommonCommentGroup" (id, name, title, "displayOrder", "isLinked", "linkedField")
         VALUES ($1, $2, $3, $4, $5, $6)`,
        [groupId, group.name, group.title, group.displayOrder, group.isLinked, group.linkedField ?? null]
      );

      for (const opt of group.options) {
        await client.query(
          `INSERT INTO "CommonCommentOption" (id, code, text, "displayOrder", "groupId")
           VALUES ($1, $2, $3, $4, $5)`,
          [createId(), opt.code, opt.text, opt.displayOrder, groupId]
        );
      }
    }

    // Seed comment format template
    const templateKey = 'comment_format_template';
    const templateValue = '<Academic>\n\n<Effort> <Behaviour> <Homework>\n\n<SCG>\n\n<Overall>';
    await client.query(
      `INSERT INTO "AppSetting" (key, value) VALUES ($1, $2)
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      [templateKey, templateValue]
    );

    console.log('Common Comment Groups seeded successfully');

    console.log('\n=== Seeding Summary ===');
    console.log(`Created ${pupilData.length} pupils`);
    console.log(`Created 4 classes: 25-7A, 25-7B, 25-7C, 25-7D`);
    console.log(`Created 3 subjects: 7CS, 7DT, 8CS`);
    console.log(`Assigned all subjects to leroysalih (HOD)`);
    console.log(`Created 6 users: admin, leroysalih, teacher, teacher2, teacher3, teacher4`);
    console.log(`Created 5 Common Comment Groups with options`);
    console.log('Seeding completed successfully');

  } finally {
    client.release();
  }
}

main()
  .catch((e) => {
    console.error(e);
    process.exit(1);
  })
  .finally(async () => {
    await pool.end();
  });
