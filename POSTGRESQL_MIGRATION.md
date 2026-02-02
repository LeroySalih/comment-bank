# PostgreSQL Migration Guide

## Current Status

The application has been updated to use PostgreSQL instead of SQLite. The Prisma
schema has been updated with:

- Database provider changed to `postgresql`
- Performance indexes added to frequently queried fields

## Required Action

Before running the migration, you need to update your `.env` file with the
PostgreSQL connection string.

### Update .env File

Replace the current `DATABASE_URL` in your `.env` file with your PostgreSQL
connection string:

```env
# Old (SQLite):
# DATABASE_URL="file:./dev.db"

# New (PostgreSQL):
DATABASE_URL="postgresql://username:password@localhost:5432/comment_bank?schema=public"
```

**Format**: `postgresql://USER:PASSWORD@HOST:PORT/DATABASE?schema=SCHEMA`

### Example Connection Strings

**Local PostgreSQL**:

```
DATABASE_URL="postgresql://postgres:password@localhost:5432/comment_bank"
```

**Docker PostgreSQL**:

```
DATABASE_URL="postgresql://postgres:password@postgres:5432/comment_bank"
```

**Cloud PostgreSQL (e.g., Supabase, Neon, Railway)**:

```
DATABASE_URL="postgresql://user:pass@db.example.com:5432/dbname?sslmode=require"
```

## Running the Migration

Once you've updated the `.env` file:

### 1. Create the Migration

```bash
npx prisma migrate dev --name migrate_to_postgresql_with_indexes
```

This will:

- Create the migration SQL file
- Apply it to your PostgreSQL database
- Generate the Prisma client

### 2. Verify the Migration

```bash
npx prisma studio
```

This opens a GUI to inspect your database.

### 3. Seed the Database (if needed)

If you have a seed file:

```bash
npx prisma db seed
```

## Data Migration from SQLite (Optional)

If you have existing data in SQLite that you want to migrate to PostgreSQL:

### Option 1: Manual Export/Import

1. Export data from SQLite:

```bash
sqlite3 prisma/dev.db .dump > backup.sql
```

2. Convert and import to PostgreSQL (requires manual SQL editing)

### Option 2: Use Prisma Studio

1. Open SQLite database in Prisma Studio
2. Export data as CSV
3. Switch to PostgreSQL
4. Import CSV data

### Option 3: Custom Migration Script

Create a Node.js script to read from SQLite and write to PostgreSQL.

## Troubleshooting

### Connection Errors

If you get connection errors:

- Verify PostgreSQL is running: `pg_isready`
- Check credentials and port
- Ensure database exists: `createdb comment_bank`

### SSL Errors

For cloud databases, you may need to add SSL parameters:

```
DATABASE_URL="postgresql://...?sslmode=require"
```

### Permission Errors

Ensure your PostgreSQL user has the necessary permissions:

```sql
GRANT ALL PRIVILEGES ON DATABASE comment_bank TO your_user;
```

## Performance Indexes Added

The migration includes the following indexes for better performance:

- `Assignment`: indexes on `classId` and `pupilId`
- `PupilCode`: indexes on `assignmentId` and `groupId`
- `Class`: index on `subjectId`
- `CommentOption`: index on `groupId`

These indexes will significantly improve query performance for:

- Loading class assignments
- Fetching pupil codes
- Filtering by subject
- Loading comment options

## Next Steps

After successful migration:

1. ✅ Test the application locally
2. ✅ Verify all data is accessible
3. ✅ Run the test suite: `npm test`
4. ✅ Build the application: `npm run build`
5. ✅ Deploy to production with PostgreSQL connection string
