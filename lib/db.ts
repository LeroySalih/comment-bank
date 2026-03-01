import { Pool } from 'pg'

function createPool() {
  return new Pool({ connectionString: process.env.DATABASE_URL! })
}

const globalForPool = globalThis as unknown as { pool: Pool }

export const pool = globalForPool.pool || createPool()

if (process.env.NODE_ENV !== 'production') globalForPool.pool = pool
