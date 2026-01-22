
const { PrismaClient } = require('@prisma/client')
const { hash } = require('bcryptjs')

const prisma = new PrismaClient()

async function main() {
  const password = await hash('password', 12)
  
  const adminRole = await prisma.role.findUnique({
    where: { name: 'admin' }
  })

  if (!adminRole) {
    throw new Error('Admin role not found. Run seed-roles.js first.')
  }

  const user = await prisma.user.upsert({
    where: { username: 'admin' },
    update: {},
    create: {
      username: 'admin',
      password,
      roles: {
        connect: { id: adminRole.id }
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
