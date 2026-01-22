
const { PrismaClient } = require('@prisma/client')
const { compare } = require('bcryptjs')
const path = require('path')
require('dotenv').config({ path: path.resolve(__dirname, '../.env') })

const prisma = new PrismaClient()

async function main() {
  console.log('DATABASE_URL:', process.env.DATABASE_URL)
  
  const user = await prisma.user.findUnique({
    where: { username: 'admin' },
    include: { roles: true }
  })

  if (!user) {
    console.log('User "admin" NOT found!')
  } else {
    console.log('User "admin" found.')
    console.log('Roles:', user.roles.map(r => r.name))
    console.log('Password Hash:', user.password)
    
    const isValid = await compare('password', user.password)
    console.log('Password "password" is valid:', isValid)
  }
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
