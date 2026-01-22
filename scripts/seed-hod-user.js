
const { PrismaClient } = require('@prisma/client')
const { hash } = require('bcryptjs')

const prisma = new PrismaClient()

async function main() {
  const password = await hash('password', 12)
  
  const hodRole = await prisma.role.findUnique({
    where: { name: 'hod' }
  })

  if (!hodRole) {
    throw new Error('HOD role not found. Run seed-roles.js first.')
  }

  const user = await prisma.user.upsert({
    where: { username: 'leroysalih' },
    update: {},
    create: {
      username: 'leroysalih',
      password,
      roles: {
        connect: { id: hodRole.id }
      }
    }
  })
  
  console.log(`Created user: ${user.username} with password: password`)
}

main()
  .then(async () => {
    await prisma.$disconnect()
  })
  .catch(async (e) => {
    console.error(e)
    await prisma.$disconnect()
    process.exit(1)
  })
