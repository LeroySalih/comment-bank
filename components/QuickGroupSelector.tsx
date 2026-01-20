'use client'

import { useState } from 'react';
import { updateStudentComment } from '@/app/actions';

type Option = {
  id: string;
  code: string;
  text: string;
};

interface QuickGroupSelectorProps {
  studentId: string;
  groupName: string;
  currentCode: string | null | undefined;
  options: Option[];
}

export default function QuickGroupSelector({ 
  studentId, 
  groupName, 
  currentCode, 
  options 
}: QuickGroupSelectorProps) {
  const [val, setVal] = useState(currentCode || "");
  const [loading, setLoading] = useState(false);

  const handleChange = async (e: React.ChangeEvent<HTMLSelectElement>) => {
    const newCode = e.target.value || null;
    setVal(newCode || "");
    setLoading(true);
    
    // If clearing (value is empty string), send null
    // If selecting, send the code
    
    // We need to find the code. wait, the value IS the code in my design below?
    // Let's use code as value.
    await updateStudentComment(studentId, groupName, newCode);
    setLoading(false);
  };

  // Sort options alphabetically by code? or generic order?
  // Usually H, M, L. Let's leave them as passed (assuming passed sorted or we sort them)
  // Options usually come unsorted from DB?
  // Let's sort simply by code length then alpha? or just alpha.
  // Actually commonly H, M, L.
  // Let's just sort by ID or insert order or Code?
  // Let's sort by code.
  const sortedOptions = [...options].sort((a,b) => a.code.localeCompare(b.code));

  return (
    <div className="relative">
        <select 
            value={val} 
            onChange={handleChange}
            disabled={loading}
            className={`block w-full rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring-blue-500 sm:text-xs py-1 px-2 ${loading ? 'opacity-50' : ''} ${val ? 'bg-blue-50 text-blue-800 font-medium' : 'text-gray-500'}`}
        >
            <option value="">-</option>
            {sortedOptions.map(opt => (
                <option key={opt.id} value={opt.code} title={opt.text}>
                    {opt.code}
                </option>
            ))}
        </select>
    </div>
  );
}
