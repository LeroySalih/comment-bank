import { prisma } from "@/lib/prisma"
import CcgManagement from "./_components/ccg-management"

export const dynamic = 'force-dynamic'

export default async function CcgPage() {
  const commonGroups = await (prisma as any).commonCommentGroup.findMany({
    orderBy: [
      { paragraphPosition: 'asc' },
      { displayOrder: 'asc' }
    ],
    include: {
      CommonCommentOption: {
        orderBy: { displayOrder: 'asc' }
      }
    }
  })

  const wrapperSetting = await (prisma as any).appSetting.findUnique({
    where: { key: 'p2_wrapper_template' }
  })

  const wrapperTemplate = wrapperSetting?.value || ''

  // Fetch available linked fields from existing assignment data
  let availableLinkedFields: string[] = []
  try {
    const result = await prisma.$queryRaw<{ field_name: string }[]>`
      SELECT DISTINCT jsonb_object_keys("linkedData") as field_name
      FROM "Assignment" WHERE "linkedData" IS NOT NULL
    `
    availableLinkedFields = result.map(r => r.field_name).sort()
  } catch {
    // No linked data yet
  }

  return (
    <div className="container mx-auto py-10">
      <div className="flex justify-between items-center mb-8">
        <div>
          <h1 className="text-3xl font-bold">Common Comment Groups</h1>
          <p className="text-gray-500 mt-1">Manage global comment groups that apply to all subjects</p>
        </div>
        <a
          href="/admin"
          className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
        >
          <span className="material-symbols-outlined text-lg">arrow_back</span>
          Back to Admin
        </a>
      </div>

      <CcgManagement groups={commonGroups} wrapperTemplate={wrapperTemplate} availableLinkedFields={availableLinkedFields} />
    </div>
  )
}
