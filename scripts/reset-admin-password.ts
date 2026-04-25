import { hash } from 'bcryptjs'
import { Pool } from 'pg'
import * as dotenv from 'dotenv'

dotenv.config()

const pool = new Pool({ connectionString: process.env.DATABASE_URL })

async function main() {
  const hashed = await hash('admin', 10)
  const result = await pool.query(
    `UPDATE "User" SET password = $1 WHERE username = 'admin' RETURNING username`,
    [hashed]
  )

  if (result.rowCount === 0) {
    console.error('No user with username "admin" found.')
    process.exit(1)
  }

  console.log('Password reset to "admin" for user:', result.rows[0].username)
  await pool.end()
}

main().catch((err) => {
  console.error(err)
  process.exit(1)
})
