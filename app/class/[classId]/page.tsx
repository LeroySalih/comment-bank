import { prisma } from '@/lib/prisma';
import Link from 'next/link';
import { ArrowLeft } from 'lucide-react';
import QuickGroupSelector from '@/components/QuickGroupSelector';

export default async function ClassPage({ params }: { params: Promise<{ classId: string }> }) {
  const { classId } = await params;

  const cls = await prisma.class.findUnique({
    where: { id: classId },
    include: {
      course: {
        include: {
            commentGroups: {
                include: { options: true }
            }
        }
      },
      students: {
        orderBy: { lastName: 'asc' }
      }
    }
  });

  if (!cls) {
    return <div>Class not found</div>;
  }

  // Sort groups: WP, TH, PS, OA
  const order = ['WP', 'TH', 'PS', 'OA'];
  const groups = cls.course.commentGroups.sort((a: any, b: any) => {
      const indexA = order.indexOf(a.name);
      const indexB = order.indexOf(b.name);
      if (indexA !== -1 && indexB !== -1) return indexA - indexB;
      if (indexA !== -1) return -1;
      if (indexB !== -1) return 1;
      return a.name.localeCompare(b.name);
  });

  return (
    <main className="min-h-screen p-8 bg-gray-50">
      <div className="max-w-6xl mx-auto">
        <header className="mb-8">
           <Link href="/" className="inline-flex items-center text-sm text-gray-500 hover:text-gray-900 mb-4 transition-colors">
            <ArrowLeft className="w-4 h-4 mr-1" />
            Back to Dashboard
           </Link>
          <div className="flex justify-between items-end">
             <div>
                <h1 className="text-3xl font-bold text-gray-900">{cls.name}</h1>
                <p className="text-gray-600">{cls.course.name}</p>
             </div>
             <div className="text-gray-500">
                {cls.students.length} Pupils
             </div>
          </div>
        </header>

        <div className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-hidden">
             <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Gender</th>
                    {groups.map((g: any) => (
                        <th key={g.id} scope="col" className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-24">
                            {g.name}
                        </th>
                    ))}
                    <th scope="col" className="relative px-6 py-3">
                      <span className="sr-only">Edit</span>
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {cls.students.map((student) => (
                    <tr key={student.id} className="hover:bg-gray-50">
                      <td className="px-6 py-4 whitespace-nowrap">
                        <div className="text-sm font-medium text-gray-900">{student.lastName}, {student.firstName}</div>
                      </td>
                      <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                        {student.gender}
                      </td>
                      {groups.map((g: any) => {
                          // Get current code for this group on this student
                          let currentCode = null;
                          if (g.name === 'WP') currentCode = student.wpCode;
                          if (g.name === 'TH') currentCode = student.thCode;
                          if (g.name === 'PS') currentCode = student.psCode;
                          if (g.name === 'OA') currentCode = student.oaCode;

                          return (
                              <td key={g.id} className="px-4 py-4 whitespace-nowrap">
                                  <QuickGroupSelector 
                                    studentId={student.id} 
                                    groupName={g.name} 
                                    currentCode={currentCode}
                                    options={g.options} 
                                  />
                              </td>
                          );
                      })}
                      <td className="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                        <Link href={`/student/${student.id}`} className="text-blue-600 hover:text-blue-900 hover:underline">
                          Write Comment
                        </Link>
                      </td>
                    </tr>
                  ))}
                </tbody>
             </table>
        </div>
      </div>
    </main>
  );
}
