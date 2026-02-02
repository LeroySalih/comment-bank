"use client"

import { useState, useEffect, useRef } from "react"
import { getPupils, updatePupil, processPupilUpload } from "@/lib/server-actions/admin"

interface Pupil {
  admissionNumber: string
  firstName: string
  lastName: string
  gender: string
  form: string | null
  isActive: boolean
}

export function PupilManagement() {
  const [query, setQuery] = useState("")
  const [pupils, setPupils] = useState<Pupil[]>([])
  const [loading, setLoading] = useState(false)
  const [uploading, setUploading] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const fetchPupils = async (q: string) => {
    setLoading(true)
    const result = await getPupils(q)
    if (result.success) {
      setPupils(result.pupils || [])
    }
    setLoading(false)
  }

  useEffect(() => {
    const timer = setTimeout(() => {
      fetchPupils(query)
    }, 300)
    return () => clearTimeout(timer)
  }, [query])

  const handleToggleActive = async (pupil: Pupil) => {
    const newStatus = !pupil.isActive
    // Optimistic update
    setPupils(pupils.map(p => 
      p.admissionNumber === pupil.admissionNumber ? { ...p, isActive: newStatus } : p
    ))

    const result = await updatePupil(pupil.admissionNumber, { isActive: newStatus })
    if (!result.success) {
      alert("Failed to update pupil status")
      // Revert
      setPupils(pupils.map(p => 
        p.admissionNumber === pupil.admissionNumber ? { ...p, isActive: pupil.isActive } : p
      ))
    }
  }

  const handleUpdateName = async (pupil: Pupil, field: 'firstName' | 'lastName', value: string) => {
    if (pupil[field] === value) return

    const result = await updatePupil(pupil.admissionNumber, { [field]: value })
    if (result.success) {
      setPupils(pupils.map(p => 
        p.admissionNumber === pupil.admissionNumber ? { ...p, [field]: value } : p
      ))
    } else {
      alert("Failed to update pupil name")
    }
  }

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = async (e) => {
      const content = e.target?.result as string
      // Extract base64 content (remove data URL prefix)
      const base64Content = content.split(',')[1]
      setUploading(true)
      const result = await processPupilUpload(base64Content)
      setUploading(false)
      
      if (result.success) {
        alert('message' in result ? result.message : 'Successfully synced pupils')
        fetchPupils(query)
      } else {
        alert('error' in result ? result.error : "Failed to process upload")
      }
    }
    reader.readAsDataURL(file)
    // Clear the input
    if (fileInputRef.current) fileInputRef.current.value = ""
  }

  return (
    <div className="space-y-6">
      <div className="bg-white shadow rounded-lg p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-xl font-semibold">Pupil Management</h2>
          <div className="flex items-center gap-4">
            <input
              type="file"
              accept=".csv,.xlsx,.xls"
              ref={fileInputRef}
              onChange={handleFileUpload}
              className="hidden"
            />
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={uploading}
              className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-md font-medium disabled:opacity-50"
            >
              {uploading ? "Uploading..." : "Upload Pupil List"}
            </button>
          </div>
        </div>

        <div className="mb-4">
          <input
            type="text"
            placeholder="Search by name or admission number..."
            className="w-full p-2 border rounded"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>

        <div className="overflow-x-auto">
          <table className="min-w-full divide-y divide-gray-200">
            <thead className="bg-gray-50">
              <tr>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Adm No</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">First Name</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Last Name</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Gender</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Form</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Status</th>
                <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="bg-white divide-y divide-gray-200">
              {loading ? (
                <tr>
                  <td colSpan={7} className="px-6 py-4 text-center text-sm text-gray-500">Loading...</td>
                </tr>
              ) : pupils.length === 0 ? (
                <tr>
                  <td colSpan={7} className="px-6 py-4 text-center text-sm text-gray-500">No pupils found</td>
                </tr>
              ) : (
                pupils.map((pupil) => (
                  <tr key={pupil.admissionNumber} className={pupil.isActive ? "" : "bg-gray-50"}>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">{pupil.admissionNumber}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                      <input
                        type="text"
                        defaultValue={pupil.firstName}
                        onBlur={(e) => handleUpdateName(pupil, 'firstName', e.target.value)}
                        className="bg-transparent border-b border-transparent focus:border-blue-500 focus:outline-none"
                      />
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                      <input
                        type="text"
                        defaultValue={pupil.lastName}
                        onBlur={(e) => handleUpdateName(pupil, 'lastName', e.target.value)}
                        className="bg-transparent border-b border-transparent focus:border-blue-500 focus:outline-none"
                      />
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{pupil.gender}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">{pupil.form || '-'}</td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-500">
                      <span className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${pupil.isActive ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
                        {pupil.isActive ? 'Active' : 'Inactive'}
                      </span>
                    </td>
                    <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                      <button
                        onClick={() => handleToggleActive(pupil)}
                        className={`${pupil.isActive ? 'text-red-600 hover:text-red-900' : 'text-green-600 hover:text-green-900'}`}
                      >
                        Set {pupil.isActive ? 'Inactive' : 'Active'}
                      </button>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
