const { prisma } = require('./lib/prisma');

async function check() {
  try {
    console.log("Checking Prisma Models...");
    // Check if we can access the field metadata via dmmf if available, 
    // or just try a query and catch the specific error details.
    
    const subject = await prisma.course.findFirst({
      include: {
        classes: {
          take: 1
        }
      }
    });
    
    if (subject && subject.classes.length > 0) {
        const clsId = subject.classes[0].id;
        console.log(`Testing query on Class ID: ${clsId}`);
        try {
            const cls = await prisma.class.findUnique({
                where: { id: clsId },
                include: { teachers: true }
            });
            console.log("Runtime check SUCCESS: teachers field found.");
        } catch (e) {
            console.log("Runtime check FAILED: teachers field NOT found.");
            console.error(e.message);
        }
    } else {
        console.log("No classes found to test.");
    }
  } catch (error) {
    console.error("Setup failed:", error.message);
  } finally {
    await prisma.$disconnect();
  }
}

check();
