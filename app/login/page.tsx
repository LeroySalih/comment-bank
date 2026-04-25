import { pool } from '@/lib/db'
import LoginForm from './login-form'

async function checkDbConnection() {
  try {
    await pool.query('SELECT 1')
    return true
  } catch (err) {
    console.error('[LOGIN] Database connection failed:', err)
    return false
  }
}

export default async function LoginPage() {
  await checkDbConnection()
  return <LoginForm />
}
