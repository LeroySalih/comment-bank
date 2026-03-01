import { pool } from "@/lib/db"
import { notFound } from "next/navigation"
import Link from "next/link"
import { PupilForm } from "./_components/pupil-form"
import { StudentRow } from "./_components/student-row"
import { decrypt } from "@/lib/encryption"

export const dynamic = 'force-dynamic'

interface Props {
  params: Promise<{
    classId: string
  }>
}

export default async function ClassPage({ params }: Props) {
  const { classId } = await params

  // Fetch class with subject
  const { rows: classRows } = await pool.query(
    `SELECT c.*, s.id as s_id, s.code as s_code, s.title as s_title
     FROM "Class" c
     JOIN "Subject" s ON s.id = c."subjectId"
     WHERE c.id = $1`,
    [classId]
  )

  if (classRows.length === 0) notFound()

  const classRow = classRows[0]
  const classData = {
    id: classRow.id,
    name: classRow.name,
    year: classRow.year,
    subjectId: classRow.subjectId,
    Subject: { id: classRow.s_id, code: classRow.s_code, title: classRow.s_title },
  }

  // Fetch assignments with pupils, sorted by decrypted lastName
  const { rows: assignmentRows } = await pool.query(
    `SELECT a.*,
            p."admissionNumber" as pupil_admissionNumber,
            p."firstName" as pupil_firstName,
            p."lastName" as pupil_lastName,
            p.gender as pupil_gender,
            p."isActive" as pupil_isActive,
            p.form as pupil_form
     FROM "Assignment" a
     JOIN "Pupil" p ON p."admissionNumber" = a."pupilId"
     WHERE a."classId" = $1`,
    [classId]
  )

  // Decrypt pupil names
  const assignments = assignmentRows.map((row: any) => ({
    id: row.id,
    pupilId: row.pupilId,
    classId: row.classId,
    eoyLevel: row.eoyLevel,
    targetLevel: row.targetLevel,
    actualLevel: row.actualLevel,
    finalComment: row.finalComment,
    linkedData: row.linkedData,
    checkStatus: row.checkStatus,
    checkNote: row.checkNote,
    checkedAt: row.checkedAt,
    checkedById: row.checkedById,
    Pupil: {
      admissionNumber: row.pupil_admissionNumber,
      firstName: decrypt(row.pupil_firstName),
      lastName: decrypt(row.pupil_lastName),
      gender: row.pupil_gender,
      isActive: row.pupil_isActive,
      form: row.pupil_form,
    }
  }))

  // Re-sort because database sort was on encrypted strings
  assignments.sort((a: any, b: any) =>
    a.Pupil.lastName.localeCompare(b.Pupil.lastName)
  )

  return (
    <div className="container mx-auto py-10">
      <div className="flex flex-col gap-2 mb-8">
        <div className="flex items-center gap-2 text-sm text-gray-500">
           <Link href="/hod" className="hover:text-indigo-600">Dashboard</Link>
           <span>/</span>
           <Link href={`/hod/subject/${classData.subjectId}`} className="hover:text-indigo-600">{classData.Subject.code}</Link>
           <span>/</span>
           <span>{classData.name}</span>
        </div>
        <h1 className="text-3xl font-bold">{classData.name} - Pupils</h1>
      </div>

      <div className="bg-white shadow rounded-lg p-6">
        <div className="flex justify-between items-center mb-6">
           <h2 className="text-xl font-semibold">Pupils</h2>
           <PupilForm classId={classData.id} />
        </div>

        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Name</th>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Gender</th>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Target Level</th>
                <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Actual Level</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {assignments.map((assignment: any) => (
                <StudentRow
                  key={assignment.id}
                  student={assignment}
                  classId={classData.id}
                />
              ))}
              {assignments.length === 0 && (
                <tr>
                  <td colSpan={4} className="px-6 py-10 text-center text-sm text-gray-500">
                    No pupils added to this class yet.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
