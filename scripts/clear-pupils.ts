import { pool } from '@/lib/db'

async function main() {
  console.log('Clearing pupils and related data...')

  // Due to CASCADE deletes in the schema, deleting Pupil handles Assignment, PupilCode, etc.
  const { rowCount } = await pool.query(`DELETE FROM "Pupil"`)

  console.log(`Successfully deleted ${rowCount} pupils.`)
}

main()
  .catch((e) => {
    console.error(e)
    process.exit(1)
  })
  .finally(async () => {
    await pool.end()
  })
