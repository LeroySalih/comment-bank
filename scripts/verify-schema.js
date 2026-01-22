const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function main() {
  console.log("Checking Class model fields...");
  // Attempt to fetch a class and check its fields dynamically
  const cls = await prisma.class.findFirst({
    include: {
      teachers: true,
      course: true,
      students: true
    }
  });

  if (cls) {
    console.log("Class found:", cls.name);
    console.log("Teachers count:", cls.teachers?.length);
    console.log("Course found:", cls.course?.name);
    console.log("Students count:", cls.students?.length);
  } else {
    console.log("No class found to test.");
  }
}

main()
  .catch(e => {
    console.error("Query failed:");
    console.error(e.message);
  })
  .finally(async () => await prisma.$disconnect());
