import { PrismaClient } from '@prisma/client'

const prisma = new PrismaClient()

async function main() {
  console.log('Clearing pupils and related data...')
  
  // Due to Cascade deletes in the schema, deleting Pupil should handle Assignment and PupilCode
  const deleteCount = await prisma.pupil.deleteMany({})
  
  console.log(`Successfully deleted ${deleteCount.count} pupils.`)
}

main()
  .catch((e) => {
    console.error(e)
    process.exit(1)
  })
  .finally(async () => {
    await prisma.$disconnect()
  })
