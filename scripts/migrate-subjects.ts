import { PrismaClient } from '@prisma/client';
import fs from 'fs';

const prisma = new PrismaClient();
const BACKUP_PATH = './data/migration_backup.json';

async function main() {
  if (!fs.existsSync(BACKUP_PATH)) {
    console.log('No backup file found, skipping migration.');
    return;
  }

  const subjects = JSON.parse(fs.readFileSync(BACKUP_PATH, 'utf-8'));
  console.log(`Found ${subjects.length} subjects to migrate.`);

  for (const sub of subjects) {
    // 1. Create/Update Subject
    const userToConnect = sub.owner ? { connect: { username: sub.owner.username } } : undefined;

    const newSubject = await prisma.subject.upsert({
      where: { code: sub.name }, // map old 'name' to 'code' as unique key
      update: {
        title: sub.subject,
        studiedComment: sub.studiedComment,
        users: userToConnect // Connect IF we have an owner
      },
      create: {
        code: sub.name,
        title: sub.subject,
        studiedComment: sub.studiedComment,
        users: userToConnect
      }
    });

    console.log(`Migrated subject: ${newSubject.code}`);

    // 2. Restore Classes
    if (sub.classes && sub.classes.length > 0) {
      for (const cls of sub.classes) {
        
        // Prepare teachers connection
        const teachersToConnect = cls.teachers && cls.teachers.length > 0
          ? cls.teachers.map((t: any) => ({ username: t.username }))
          : [];

        // Connect teachers to Subject as well, as per new rule?
        // No, user requirement: "Subject-level assignment (Admin only) grants access to all classes"
        // But we ALSO kept Class-level assignment.
        // So we restore the granular class assignment.
        
        // Also ensure teachers are linked to the subject?
        // "An admin can assign Admin, HoD and Teachers to a subject"
        // If a teacher was assigned to a class, should they now be assigned to the subject?
        // Probably NOT automatically, otherwise they get access to ALL classes.
        // Keep them on the class only.

        await prisma.class.upsert({
          where: { name: cls.name },
          update: {
            year: cls.year,
            subjectId: newSubject.id,
            teachers: { connect: teachersToConnect }
          },
          create: {
            name: cls.name,
            year: cls.year,
            subjectId: newSubject.id,
            teachers: { connect: teachersToConnect }
          }
        });
      }
    }
  }

  console.log('Migration completed.');
}

main()
  .catch(e => console.error(e))
  .finally(() => prisma.$disconnect());
