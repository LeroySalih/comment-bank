"use client"

import { parseComment } from "@/lib/utils"

interface VariablePreviewProps {
  text: string
  subjectName?: string | null
}

export function VariablePreview({ text, subjectName }: VariablePreviewProps) {
  if (!text) return null

  const malePreview = parseComment(
    text,
    "Alex",
    "Male",
    subjectName || "Computer Science",
    "11",
    "8",
    "9"
  )

  const femalePreview = parseComment(
    text,
    "Alexandra",
    "Female",
    subjectName || "Computer Science",
    "11",
    "8",
    "9"
  )

  return (
    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mt-2 mb-4 p-3 bg-indigo-50/50 rounded-lg border border-indigo-100">
      <div>
        <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-wider">Male Preview</span>
        <p className="text-sm text-gray-700 mt-1 leading-relaxed">
          {malePreview || <span className="italic text-gray-400">Preview will appear here...</span>}
        </p>
      </div>
      <div className="border-t md:border-t-0 md:border-l border-indigo-100 pt-3 md:pt-0 md:pl-4">
        <span className="text-[10px] font-bold text-pink-400 uppercase tracking-wider">Female Preview</span>
        <p className="text-sm text-gray-700 mt-1 leading-relaxed">
          {femalePreview || <span className="italic text-gray-400">Preview will appear here...</span>}
        </p>
      </div>
      <div className="col-span-full pt-2 border-t border-indigo-50">
        <p className="text-[10px] text-gray-400 italic">
          Variables: {'<Name>'}, {'<he/she>'}, {'<his/her>'}, {'<him/her>'}, {'<Subject>'}, {'<Year>'}, {'<EoYLevel>'}, {'<TargetLevel>'}
        </p>
      </div>
    </div>
  )
}
