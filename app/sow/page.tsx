import { getServerSession } from 'next-auth';
import { authOptions } from '@/app/api/auth/[...nextauth]/route';
import { pool } from '@/lib/db';
import Link from 'next/link';
import { BookOpen } from 'lucide-react';
import { redirect } from 'next/navigation';

export default async function SowIndexPage() {
  const session = await getServerSession(authOptions);
  if (!session) redirect('/login');

  // Show classes the logged-in teacher teaches
  const classes = await pool.query<{ name: string; subjectTitle: string }>(
    `SELECT c.name, s.title as "subjectTitle"
     FROM "Class" c
     JOIN "Subject" s ON s.id = c."subjectId"
     JOIN "_ClassToUser" cu ON cu."A" = c.id
     JOIN "User" u ON u.id = cu."B"
     WHERE u.username = $1
     ORDER BY c.name`,
    [session.user?.username],
  );

  return (
    <div className="max-w-4xl mx-auto px-4 py-10">
      <div className="flex items-center gap-3 mb-8">
        <BookOpen className="text-blue-600" size={28} />
        <h1 className="text-2xl font-bold text-gray-900">Schemes of Work</h1>
      </div>

      {classes.rows.length === 0 ? (
        <p className="text-gray-500 text-sm">You have no classes assigned.</p>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
          {classes.rows.map((cls) => (
            <Link
              key={cls.name}
              href={`/sow/${encodeURIComponent(cls.name)}`}
              className="block border border-gray-200 rounded-xl p-5 hover:border-blue-400 hover:shadow-sm transition-all"
            >
              <p className="text-xs font-mono text-gray-400 mb-1">{cls.name}</p>
              <p className="font-semibold text-gray-800">{cls.subjectTitle ?? cls.name}</p>
            </Link>
          ))}
        </div>
      )}
    </div>
  );
}
