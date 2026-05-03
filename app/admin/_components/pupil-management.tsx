"use client"

import React, { useState, useEffect, useRef } from "react"
import { getPupils, updatePupil, processPupilUpload, createPupil, deletePupil, getClassesForPupilAssignment, addPupilsToClass } from "@/lib/server-actions/admin"

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
  const [showAddForm, setShowAddForm] = useState(false)
  const [addForm, setAddForm] = useState({ admissionNumber: '', firstName: '', lastName: '', gender: 'M', form: '' })
  const [addError, setAddError] = useState<string | null>(null)
  const [addSaving, setAddSaving] = useState(false)
  const [deletingId, setDeletingId] = useState<string | null>(null)
  const [assigningPupilId, setAssigningPupilId] = useState<string | null>(null)
  const [assignClasses, setAssignClasses] = useState<Array<{
    id: string
    name: string
    year: string | null
    subjectTitle: string
    isAssigned: boolean
  }>>([])
  const [assignLoading, setAssignLoading] = useState(false)
  const [assigningClassId, setAssigningClassId] = useState<string | null>(null)

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

  const handleAddPupil = async (e: React.FormEvent) => {
    e.preventDefault()
    setAddError(null)
    setAddSaving(true)
    const result = await createPupil({
      admissionNumber: addForm.admissionNumber.trim(),
      firstName: addForm.firstName.trim(),
      lastName: addForm.lastName.trim(),
      gender: addForm.gender as 'M' | 'F',
      form: addForm.form.trim() || null
    })
    setAddSaving(false)
    if (result.success) {
      setAddForm({ admissionNumber: '', firstName: '', lastName: '', gender: 'M', form: '' })
      setShowAddForm(false)
      fetchPupils(query)
    } else {
      setAddError('error' in result ? (result.error ?? 'Failed to create pupil') : 'Failed to create pupil')
    }
  }

  const handleDeletePupil = async (pupil: Pupil) => {
    if (!confirm(`Permanently delete ${pupil.firstName} ${pupil.lastName} (${pupil.admissionNumber})? This will also remove them from all classes.`)) return
    setDeletingId(pupil.admissionNumber)
    const result = await deletePupil(pupil.admissionNumber)
    if (result.success) {
      setPupils(prev => prev.filter(p => p.admissionNumber !== pupil.admissionNumber))
    } else {
      alert('error' in result ? (result.error ?? 'Failed to delete pupil') : 'Failed to delete pupil')
    }
    setDeletingId(null)
  }

  const openAssignPanel = async (admissionNumber: string) => {
    if (assigningPupilId === admissionNumber) {
      setAssigningPupilId(null)
      return
    }
    setAssigningPupilId(admissionNumber)
    setAssignLoading(true)
    const result = await getClassesForPupilAssignment(admissionNumber)
    if (result.success && 'classes' in result) {
      setAssignClasses(result.classes)
    }
    setAssignLoading(false)
  }

  const handleAssignToClass = async (admissionNumber: string, classId: string) => {
    setAssigningClassId(classId)
    const result = await addPupilsToClass(classId, [admissionNumber])
    if (result.success) {
      setAssignClasses(prev => prev.map(c => c.id === classId ? { ...c, isAssigned: true } : c))
    } else {
      alert('error' in result ? (result.error ?? 'Failed to assign to class') : 'Failed to assign')
    }
    setAssigningClassId(null)
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
              type="button"
              onClick={() => { setShowAddForm(v => !v); setAddError(null) }}
              className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-md font-medium"
            >
              {showAddForm ? 'Cancel' : '+ Add Pupil'}
            </button>
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={uploading}
              className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-md font-medium disabled:opacity-50"
            >
              {uploading ? "Uploading..." : "Upload Pupil List"}
            </button>
          </div>
        </div>

        {showAddForm && (
          <form onSubmit={handleAddPupil} className="mb-4 p-4 bg-green-50 border border-green-200 rounded-lg space-y-3">
            <h3 className="text-sm font-semibold text-green-800">Add New Pupil</h3>
            <div className="flex flex-wrap gap-3">
              <div className="w-36">
                <label className="block text-xs font-medium text-gray-600 mb-1">Admission No *</label>
                <input
                  type="text"
                  required
                  value={addForm.admissionNumber}
                  onChange={e => setAddForm(f => ({ ...f, admissionNumber: e.target.value }))}
                  placeholder="e.g. 12345"
                  className="w-full border rounded px-2 py-1.5 text-sm"
                />
              </div>
              <div className="w-36">
                <label className="block text-xs font-medium text-gray-600 mb-1">First Name *</label>
                <input
                  type="text"
                  required
                  value={addForm.firstName}
                  onChange={e => setAddForm(f => ({ ...f, firstName: e.target.value }))}
                  className="w-full border rounded px-2 py-1.5 text-sm"
                />
              </div>
              <div className="w-36">
                <label className="block text-xs font-medium text-gray-600 mb-1">Last Name *</label>
                <input
                  type="text"
                  required
                  value={addForm.lastName}
                  onChange={e => setAddForm(f => ({ ...f, lastName: e.target.value }))}
                  className="w-full border rounded px-2 py-1.5 text-sm"
                />
              </div>
              <div className="w-28">
                <label className="block text-xs font-medium text-gray-600 mb-1">Gender *</label>
                <select
                  value={addForm.gender}
                  onChange={e => setAddForm(f => ({ ...f, gender: e.target.value }))}
                  className="w-full border rounded px-2 py-1.5 text-sm"
                >
                  <option value="M">M</option>
                  <option value="F">F</option>
                </select>
              </div>
              <div className="w-28">
                <label className="block text-xs font-medium text-gray-600 mb-1">Form</label>
                <input
                  type="text"
                  value={addForm.form}
                  onChange={e => setAddForm(f => ({ ...f, form: e.target.value }))}
                  placeholder="e.g. 9A"
                  className="w-full border rounded px-2 py-1.5 text-sm"
                />
              </div>
            </div>
            {addError && <p className="text-red-600 text-xs">{addError}</p>}
            <div className="flex gap-2">
              <button
                type="submit"
                disabled={addSaving}
                className="bg-green-600 hover:bg-green-700 text-white px-4 py-1.5 rounded text-sm font-medium disabled:opacity-50"
              >
                {addSaving ? 'Creating...' : 'Create Pupil'}
              </button>
              <button
                type="button"
                onClick={() => { setShowAddForm(false); setAddError(null) }}
                className="bg-gray-200 hover:bg-gray-300 text-gray-700 px-4 py-1.5 rounded text-sm"
              >
                Cancel
              </button>
            </div>
          </form>
        )}

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
                  <td colSpan={8} className="px-6 py-4 text-center text-sm text-gray-500">Loading...</td>
                </tr>
              ) : pupils.length === 0 ? (
                <tr>
                  <td colSpan={8} className="px-6 py-4 text-center text-sm text-gray-500">No pupils found</td>
                </tr>
              ) : (
                pupils.map((pupil) => (
                  <React.Fragment key={pupil.admissionNumber}>
                  <tr className={pupil.isActive ? "" : "bg-gray-50"}>
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
                      <div className="flex items-center gap-3">
                        <button
                          type="button"
                          onClick={() => handleToggleActive(pupil)}
                          className={`${pupil.isActive ? 'text-red-600 hover:text-red-900' : 'text-green-600 hover:text-green-900'}`}
                        >
                          Set {pupil.isActive ? 'Inactive' : 'Active'}
                        </button>
                        <button
                          type="button"
                          onClick={() => handleDeletePupil(pupil)}
                          disabled={deletingId === pupil.admissionNumber}
                          className="text-gray-400 hover:text-red-600 disabled:opacity-50 transition-colors"
                          title="Delete pupil permanently"
                        >
                          {deletingId === pupil.admissionNumber ? '…' : '🗑'}
                        </button>
                        <button
                          type="button"
                          onClick={() => openAssignPanel(pupil.admissionNumber)}
                          className={`text-sm font-medium transition-colors ${assigningPupilId === pupil.admissionNumber ? 'text-blue-700' : 'text-blue-500 hover:text-blue-800'}`}
                          title="Assign to a class"
                        >
                          + Class
                        </button>
                      </div>
                    </td>
                  </tr>
                  {assigningPupilId === pupil.admissionNumber && (
                    <tr className="bg-blue-50">
                      <td colSpan={8} className="px-6 py-3">
                        <div className="text-xs font-semibold text-blue-700 mb-2">Assign to a class</div>
                        {assignLoading ? (
                          <p className="text-xs text-gray-400">Loading classes…</p>
                        ) : (
                          <div className="flex flex-wrap gap-2">
                            {assignClasses.filter(c => !c.isAssigned).length === 0 ? (
                              <p className="text-xs text-gray-500">This pupil is already assigned to all classes.</p>
                            ) : (
                              assignClasses.filter(c => !c.isAssigned).map(c => (
                                <button
                                  key={c.id}
                                  type="button"
                                  onClick={() => handleAssignToClass(pupil.admissionNumber, c.id)}
                                  disabled={assigningClassId === c.id}
                                  className="flex items-center gap-1 px-2 py-1 bg-white border border-blue-200 rounded text-xs hover:bg-blue-100 disabled:opacity-50 transition-colors"
                                >
                                  <span className="font-semibold text-blue-700">{c.name}</span>
                                  <span className="text-gray-500">— {c.subjectTitle}{c.year ? ` (Y${c.year})` : ''}</span>
                                  {assigningClassId === c.id ? <span className="text-blue-400 ml-1">…</span> : <span className="text-blue-400 ml-1">+</span>}
                                </button>
                              ))
                            )}
                            {assignClasses.filter(c => c.isAssigned).length > 0 && (
                              <div className="w-full mt-1">
                                <span className="text-xs text-gray-400">
                                  Already in: {assignClasses.filter(c => c.isAssigned).map(c => c.name).join(', ')}
                                </span>
                              </div>
                            )}
                          </div>
                        )}
                      </td>
                    </tr>
                  )}
                  </React.Fragment>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}
