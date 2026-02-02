"use client"

import { useState, useEffect } from "react"
import { Search, UserPlus, X, Filter } from "lucide-react"
import { getClassPupils, getAllForms, getPupils, addPupilsToClass, removePupilsFromClass } from "@/lib/server-actions/admin"

interface Pupil {
  admissionNumber: string
  firstName: string
  lastName: string
  gender: string
  form: string | null
  isActive: boolean
  assignmentId?: string
}

interface ClassPupilManagerProps {
  classId: string
  onPupilsChanged?: () => void
}

export function ClassPupilManager({ classId, onPupilsChanged }: ClassPupilManagerProps) {
  const [pupils, setPupils] = useState<Pupil[]>([])
  const [allPupils, setAllPupils] = useState<Pupil[]>([])
  const [forms, setForms] = useState<string[]>([])
  const [loading, setLoading] = useState(true)
  const [filterForm, setFilterForm] = useState<string>("")
  const [searchQuery, setSearchQuery] = useState("")
  const [showAddPanel, setShowAddPanel] = useState(false)
  const [addFilterForm, setAddFilterForm] = useState<string>("")
  const [addSearchQuery, setAddSearchQuery] = useState("")
  const [selectedToAdd, setSelectedToAdd] = useState<string[]>([])
  const [selectedToRemove, setSelectedToRemove] = useState<string[]>([])
  const [saving, setSaving] = useState(false)

  const fetchData = async () => {
    setLoading(true)
    const [pupilsResult, formsResult] = await Promise.all([
      getClassPupils(classId),
      getAllForms()
    ])

    if (pupilsResult.success && pupilsResult.pupils) {
      setPupils(pupilsResult.pupils)
    }
    if (formsResult.success && formsResult.forms) {
      setForms(formsResult.forms)
    }
    setLoading(false)
  }

  const fetchAllPupils = async () => {
    const result = await getPupils("")
    if (result.success && result.pupils) {
      setAllPupils(result.pupils)
    }
  }

  useEffect(() => {
    fetchData()
  }, [classId])

  useEffect(() => {
    if (showAddPanel) {
      fetchAllPupils()
    }
  }, [showAddPanel])

  // Filter current class pupils
  const filteredPupils = pupils.filter(p => {
    const matchesForm = !filterForm || p.form === filterForm
    const matchesSearch = !searchQuery ||
      p.firstName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      p.lastName.toLowerCase().includes(searchQuery.toLowerCase()) ||
      p.admissionNumber.includes(searchQuery)
    return matchesForm && matchesSearch
  })

  // Filter available pupils (not already in class)
  const currentPupilIds = new Set(pupils.map(p => p.admissionNumber))
  const availablePupils = allPupils.filter(p => {
    if (currentPupilIds.has(p.admissionNumber)) return false
    if (!p.isActive) return false
    const matchesForm = !addFilterForm || p.form === addFilterForm
    const matchesSearch = !addSearchQuery ||
      p.firstName.toLowerCase().includes(addSearchQuery.toLowerCase()) ||
      p.lastName.toLowerCase().includes(addSearchQuery.toLowerCase()) ||
      p.admissionNumber.includes(addSearchQuery)
    return matchesForm && matchesSearch
  })

  const handleAddPupils = async () => {
    if (selectedToAdd.length === 0) return
    setSaving(true)
    const result = await addPupilsToClass(classId, selectedToAdd)
    setSaving(false)
    if (result.success) {
      setSelectedToAdd([])
      setShowAddPanel(false)
      fetchData()
      onPupilsChanged?.()
    }
  }

  const handleRemovePupils = async () => {
    if (selectedToRemove.length === 0) return
    if (!confirm(`Remove ${selectedToRemove.length} pupil(s) from this class?`)) return
    setSaving(true)
    const result = await removePupilsFromClass(classId, selectedToRemove)
    setSaving(false)
    if (result.success) {
      setSelectedToRemove([])
      fetchData()
      onPupilsChanged?.()
    }
  }

  const toggleAddSelection = (admNo: string) => {
    setSelectedToAdd(prev =>
      prev.includes(admNo) ? prev.filter(id => id !== admNo) : [...prev, admNo]
    )
  }

  const toggleRemoveSelection = (admNo: string) => {
    setSelectedToRemove(prev =>
      prev.includes(admNo) ? prev.filter(id => id !== admNo) : [...prev, admNo]
    )
  }

  const selectAllAvailable = () => {
    setSelectedToAdd(availablePupils.map(p => p.admissionNumber))
  }

  if (loading) {
    return <div className="p-4 text-center text-gray-500">Loading pupils...</div>
  }

  return (
    <div className="border-t pt-4 mt-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="text-sm font-semibold text-gray-700">
          Class Pupils ({pupils.length})
        </h3>
        <button
          type="button"
          onClick={() => setShowAddPanel(!showAddPanel)}
          className="flex items-center gap-1 text-sm text-blue-600 hover:text-blue-800"
        >
          {showAddPanel ? <X size={16} /> : <UserPlus size={16} />}
          {showAddPanel ? "Cancel" : "Add Pupils"}
        </button>
      </div>

      {/* Add Pupils Panel */}
      {showAddPanel && (
        <div className="mb-4 p-3 bg-blue-50 rounded-lg border border-blue-200">
          <h4 className="text-sm font-medium text-blue-800 mb-2">Add Pupils to Class</h4>

          {/* Filters */}
          <div className="flex gap-2 mb-2">
            <div className="relative flex-1">
              <Search size={14} className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" />
              <input
                type="text"
                value={addSearchQuery}
                onChange={(e) => setAddSearchQuery(e.target.value)}
                placeholder="Search..."
                className="w-full pl-7 pr-2 py-1.5 text-sm border rounded focus:ring-1 focus:ring-blue-500"
              />
            </div>
            <select
              value={addFilterForm}
              onChange={(e) => setAddFilterForm(e.target.value)}
              className="px-2 py-1.5 text-sm border rounded focus:ring-1 focus:ring-blue-500"
            >
              <option value="">All Forms</option>
              {forms.map(form => (
                <option key={form} value={form}>{form}</option>
              ))}
            </select>
          </div>

          {/* Available pupils list */}
          <div className="max-h-40 overflow-y-auto bg-white rounded border">
            {availablePupils.length === 0 ? (
              <p className="p-2 text-sm text-gray-500 text-center">No pupils available</p>
            ) : (
              <>
                <button
                  type="button"
                  onClick={selectAllAvailable}
                  className="w-full p-1.5 text-xs text-blue-600 hover:bg-blue-50 border-b"
                >
                  Select all ({availablePupils.length})
                </button>
                {availablePupils.map(pupil => (
                  <label
                    key={pupil.admissionNumber}
                    className={`flex items-center gap-2 p-2 text-sm hover:bg-gray-50 cursor-pointer border-b last:border-b-0 ${
                      selectedToAdd.includes(pupil.admissionNumber) ? "bg-blue-50" : ""
                    }`}
                  >
                    <input
                      type="checkbox"
                      checked={selectedToAdd.includes(pupil.admissionNumber)}
                      onChange={() => toggleAddSelection(pupil.admissionNumber)}
                      className="h-3.5 w-3.5 text-blue-600 rounded"
                    />
                    <span className="flex-1">
                      {pupil.lastName}, {pupil.firstName}
                    </span>
                    <span className="text-xs text-gray-400">{pupil.form || "-"}</span>
                  </label>
                ))}
              </>
            )}
          </div>

          {/* Add button */}
          {selectedToAdd.length > 0 && (
            <button
              type="button"
              onClick={handleAddPupils}
              disabled={saving}
              className="mt-2 w-full py-1.5 bg-blue-600 text-white text-sm rounded hover:bg-blue-700 disabled:opacity-50"
            >
              {saving ? "Adding..." : `Add ${selectedToAdd.length} Pupil(s)`}
            </button>
          )}
        </div>
      )}

      {/* Current Pupils Filter */}
      <div className="flex gap-2 mb-2">
        <div className="relative flex-1">
          <Search size={14} className="absolute left-2 top-1/2 -translate-y-1/2 text-gray-400" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search pupils..."
            className="w-full pl-7 pr-2 py-1.5 text-sm border rounded focus:ring-1 focus:ring-blue-500"
          />
        </div>
        <select
          value={filterForm}
          onChange={(e) => setFilterForm(e.target.value)}
          className="px-2 py-1.5 text-sm border rounded focus:ring-1 focus:ring-blue-500"
        >
          <option value="">All Forms</option>
          {forms.map(form => (
            <option key={form} value={form}>{form}</option>
          ))}
        </select>
      </div>

      {/* Current Pupils List */}
      <div className="max-h-48 overflow-y-auto border rounded">
        {filteredPupils.length === 0 ? (
          <p className="p-3 text-sm text-gray-500 text-center">
            {pupils.length === 0 ? "No pupils in this class" : "No pupils match the filter"}
          </p>
        ) : (
          filteredPupils.map(pupil => (
            <label
              key={pupil.admissionNumber}
              className={`flex items-center gap-2 p-2 text-sm hover:bg-gray-50 cursor-pointer border-b last:border-b-0 ${
                selectedToRemove.includes(pupil.admissionNumber) ? "bg-red-50" : ""
              }`}
            >
              <input
                type="checkbox"
                checked={selectedToRemove.includes(pupil.admissionNumber)}
                onChange={() => toggleRemoveSelection(pupil.admissionNumber)}
                className="h-3.5 w-3.5 text-red-600 rounded"
              />
              <span className="flex-1">
                {pupil.lastName}, {pupil.firstName}
              </span>
              <span className="text-xs text-gray-400">{pupil.form || "-"}</span>
            </label>
          ))
        )}
      </div>

      {/* Remove button */}
      {selectedToRemove.length > 0 && (
        <button
          type="button"
          onClick={handleRemovePupils}
          disabled={saving}
          className="mt-2 w-full py-1.5 bg-red-600 text-white text-sm rounded hover:bg-red-700 disabled:opacity-50"
        >
          {saving ? "Removing..." : `Remove ${selectedToRemove.length} Pupil(s)`}
        </button>
      )}
    </div>
  )
}
