const { PrismaClient } = require('@prisma/client')
const prisma = new PrismaClient()

const FIRST_NAMES = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda", "William", "Elizabeth", "David", "Barbara", "Richard", "Susan", "Joseph", "Jessica", "Thomas", "Sarah", "Charles", "Karen"]
const LAST_NAMES = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin"]

function getRandomElement(arr) {
  return arr[Math.floor(Math.random() * arr.length)]
}

function getRandomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

async function main() {
  console.log('Start seeding test data...')

  // 1. Fetch all courses
  const courses = await prisma.course.findMany()
  console.log(`Found ${courses.length} courses.`)

  // 2. Create a Test Class for each course
  for (const course of courses) {
    const className = `Test Class ${course.name} ${getRandomInt(100, 999)}`
    await prisma.class.create({
      data: {
        name: className,
        courseId: course.id,
        year: "Test"
      }
    })
    console.log(`Created class: ${className}`)
  }

  // 3. Fetch all classes (now including the new ones)
  const classes = await prisma.class.findMany()
  console.log(`Found ${classes.length} classes total.`)

  // 4. Create 5 random pupils for each class
  for (const cls of classes) {
    const pupilCount = await prisma.student.count({ where: { classId: cls.id } })
    if (pupilCount >= 5) {
      console.log(`Class ${cls.name} already has ${pupilCount} pupils, skipping.`)
      continue
    }

    const toCreate = 5 - pupilCount
    console.log(`Adding ${toCreate} pupils to ${cls.name}...`)

    for (let i = 0; i < toCreate; i++) {
        await prisma.student.create({
            data: {
                firstName: getRandomElement(FIRST_NAMES),
                lastName: getRandomElement(LAST_NAMES),
                gender: getRandomElement(["Male", "Female"]),
                targetLevel: String(getRandomInt(1, 9)),
                actualLevel: String(getRandomInt(1, 9)),
                classId: cls.id
            }
        })
    }
  }

  console.log('Seeding finished.')
}

main()
  .catch((e) => {
    console.error(e)
    process.exit(1)
  })
  .finally(async () => {
    await prisma.$disconnect()
  })
