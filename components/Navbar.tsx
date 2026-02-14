import Link from 'next/link';
import { getServerSession } from 'next-auth';
import { authOptions } from '../app/api/auth/[...nextauth]/route';
import { isAdmin, isHoD } from '../lib/access-control';
import SignOutButton from './SignOutButton';
import { LayoutDashboard, ShieldCheck, GraduationCap } from 'lucide-react';

export default async function Navbar() {
  const session = await getServerSession(authOptions);

  if (!session) return null;

  const showAdmin = isAdmin(session.user);
  const showHoD = isHoD(session.user);

  const envLabel = process.env.NEXT_PUBLIC_ENV || 'DEV';
  const envColor = envLabel === 'PROD' ? 'bg-red-600' : envLabel === 'TEST' ? 'bg-amber-500' : 'bg-green-600';

  return (
    <nav className="bg-white border-b border-gray-200 shadow-sm sticky top-0 z-50">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex justify-between h-16">
          <div className="flex items-center">
            <Link
              href="/"
              className="flex-shrink-0 flex items-center gap-2 text-xl font-bold text-blue-600 hover:text-blue-700 transition-colors"
            >
              <div className="bg-blue-600 text-white p-1.5 rounded-lg">
                <GraduationCap size={24} />
              </div>
              <span className="hidden sm:block">Comment Bank</span>
            </Link>
            <span className={`ml-2 px-2 py-0.5 text-[10px] font-bold text-white rounded ${envColor}`}>
              {envLabel}
            </span>
            
            <div className="hidden sm:ml-8 sm:flex sm:space-x-4">
              <Link
                href="/"
                className="inline-flex items-center px-3 py-2 text-sm font-medium text-gray-700 hover:text-blue-600 hover:bg-blue-50 rounded-md transition-all gap-2"
              >
                <LayoutDashboard size={18} />
                Pupil Comments
              </Link>

              {showHoD && (
                <Link
                  href="/hod"
                  className="inline-flex items-center px-3 py-2 text-sm font-medium text-gray-700 hover:text-blue-600 hover:bg-blue-50 rounded-md transition-all gap-2"
                >
                  <GraduationCap size={18} />
                  HOD
                </Link>
              )}

              {showAdmin && (
                <Link
                  href="/admin"
                  className="inline-flex items-center px-3 py-2 text-sm font-medium text-gray-700 hover:text-blue-600 hover:bg-blue-50 rounded-md transition-all gap-2"
                >
                  <ShieldCheck size={18} />
                  Admin
                </Link>
              )}
            </div>
          </div>
          
          <div className="flex items-center gap-4">
            <div className="hidden md:flex flex-col items-end mr-2 text-xs">
              <span className="font-semibold text-gray-900">{session.user?.username || session.user?.name}</span>
              <span className="text-gray-500 capitalize">{session.user?.roles?.[0] || 'User'}</span>
            </div>
            <SignOutButton />
          </div>
        </div>
      </div>
    </nav>
  );
}
