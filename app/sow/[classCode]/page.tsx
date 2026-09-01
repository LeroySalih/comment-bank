'use client';

import { useEffect, useState, useCallback } from 'react';
import { useParams } from 'next/navigation';
import { Plus, X, ChevronDown } from 'lucide-react';

type HalfTerm = {
  id: string;
  label: string;
  startDate: string;
  endDate: string;
};

type SowUnit = {
  id: string;
  halfTermId: string;
  title: string;
  comment: string | null;
  isManual: boolean;
  hasLessons: boolean;
};

type SowData = {
  cls: { id: string; name: string; subjectTitle: string };
  academicYear: string;
  halfTerms: HalfTerm[];
  units: SowUnit[];
};

const ACADEMIC_YEARS = ['2024/25', '2025/26', '2026/27', '2027/28'];

function formatDateRange(start: string, end: string) {
  const fmt = (d: string) => {
    const dt = new Date(d);
    return dt.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
  };
  return `${fmt(start)} – ${fmt(end)}`;
}

export default function SowPage() {
  const { classCode } = useParams<{ classCode: string }>();
  const [academicYear, setAcademicYear] = useState(() => currentAcademicYear());
  const [data, setData] = useState<SowData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Adding a unit: which half-term is being targeted
  const [addingToHalfTerm, setAddingToHalfTerm] = useState<string | null>(null);
  const [newUnitTitle, setNewUnitTitle] = useState('');

  // Editing a unit comment
  const [editingUnitId, setEditingUnitId] = useState<string | null>(null);
  const [editComment, setEditComment] = useState('');

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/sow/${classCode}?academicYear=${encodeURIComponent(academicYear)}`);
      if (!res.ok) throw new Error(await res.text());
      setData(await res.json());
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to load');
    } finally {
      setLoading(false);
    }
  }, [classCode, academicYear]);

  useEffect(() => { load(); }, [load]);

  const unitsForHalfTerm = (htId: string) =>
    (data?.units ?? []).filter((u) => u.halfTermId === htId);

  async function handleAddUnit(halfTermId: string) {
    if (!newUnitTitle.trim()) return;
    await fetch(`/api/sow/${classCode}/units`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ halfTermId, title: newUnitTitle.trim() }),
    });
    setAddingToHalfTerm(null);
    setNewUnitTitle('');
    load();
  }

  async function handleSaveComment(unitId: string) {
    await fetch(`/api/sow/${classCode}/units/${unitId}`, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ comment: editComment }),
    });
    setEditingUnitId(null);
    load();
  }

  async function handleDeleteUnit(unitId: string) {
    const res = await fetch(`/api/sow/${classCode}/units/${unitId}`, { method: 'DELETE' });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      alert(body.error ?? 'Could not delete unit');
      return;
    }
    load();
  }

  function openEdit(unit: SowUnit) {
    setEditingUnitId(unit.id);
    setEditComment(unit.comment ?? '');
  }

  if (loading) return <div className="p-8 text-gray-500">Loading…</div>;
  if (error) return <div className="p-8 text-red-600">{error}</div>;
  if (!data) return null;

  const { cls, halfTerms } = data;

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
      {/* Header */}
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-xl font-semibold text-gray-900">
          <span className="text-gray-500 font-normal">{cls.name} · </span>
          {cls.subjectTitle} — Scheme of Work
        </h1>

        {/* Academic year picker */}
        <div className="relative">
          <select
            value={academicYear}
            onChange={(e) => setAcademicYear(e.target.value)}
            className="appearance-none pl-3 pr-8 py-1.5 text-sm border border-gray-300 rounded-md bg-white text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            {ACADEMIC_YEARS.map((y) => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>
          <ChevronDown size={14} className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 pointer-events-none" />
        </div>
      </div>

      {/* Half-term grid */}
      {halfTerms.length === 0 ? (
        <div className="border border-gray-200 rounded-lg p-6 text-center text-gray-400 text-sm mb-6">
          No half-terms configured for {academicYear}.
        </div>
      ) : (
        <div className="border border-gray-200 rounded-lg overflow-hidden mb-8">
          {/* Column headers */}
          <div
            className="grid border-b border-gray-200"
            style={{ gridTemplateColumns: `repeat(${halfTerms.length}, minmax(0, 1fr))` }}
          >
            {halfTerms.map((ht, i) => (
              <div
                key={ht.id}
                className={`px-3 py-2 ${i < halfTerms.length - 1 ? 'border-r border-gray-200' : ''}`}
              >
                <p className="text-sm font-semibold text-gray-800">{ht.label}</p>
                <p className="text-xs text-gray-500">{formatDateRange(ht.startDate, ht.endDate)}</p>
              </div>
            ))}
          </div>

          {/* Unit cells */}
          <div
            className="grid"
            style={{ gridTemplateColumns: `repeat(${halfTerms.length}, minmax(0, 1fr))` }}
          >
            {halfTerms.map((ht, i) => {
              const units = unitsForHalfTerm(ht.id);
              const isAdding = addingToHalfTerm === ht.id;

              return (
                <div
                  key={ht.id}
                  className={`p-2 min-h-[80px] ${i < halfTerms.length - 1 ? 'border-r border-gray-200' : ''}`}
                >
                  {/* Existing units */}
                  {units.map((unit) => (
                    <div
                      key={unit.id}
                      className={`group relative mb-1.5 rounded px-2 py-1 text-xs cursor-pointer ${
                        unit.hasLessons
                          ? 'bg-green-100 text-green-800 hover:bg-green-200'
                          : 'bg-gray-100 text-gray-600 hover:bg-gray-200'
                      }`}
                      onClick={() => openEdit(unit)}
                    >
                      <span className="font-medium block truncate pr-4">{unit.title}</span>
                      {unit.comment && (
                        <span className="block truncate text-[10px] opacity-70 mt-0.5">{unit.comment}</span>
                      )}

                      {/* Delete button — only for manual units with no lessons */}
                      {unit.isManual && !unit.hasLessons && (
                        <button
                          onClick={(e) => { e.stopPropagation(); handleDeleteUnit(unit.id); }}
                          className="absolute top-0.5 right-0.5 hidden group-hover:flex items-center justify-center w-4 h-4 rounded text-gray-400 hover:text-red-500"
                          title="Remove unit"
                        >
                          <X size={10} />
                        </button>
                      )}
                    </div>
                  ))}

                  {/* Add unit inline */}
                  {isAdding ? (
                    <div className="mt-1">
                      <input
                        autoFocus
                        value={newUnitTitle}
                        onChange={(e) => setNewUnitTitle(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') handleAddUnit(ht.id);
                          if (e.key === 'Escape') { setAddingToHalfTerm(null); setNewUnitTitle(''); }
                        }}
                        placeholder="Unit name…"
                        className="w-full text-xs border border-blue-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-blue-500"
                      />
                      <div className="flex gap-1 mt-1">
                        <button
                          onClick={() => handleAddUnit(ht.id)}
                          className="text-[10px] px-2 py-0.5 bg-blue-600 text-white rounded hover:bg-blue-700"
                        >
                          Add
                        </button>
                        <button
                          onClick={() => { setAddingToHalfTerm(null); setNewUnitTitle(''); }}
                          className="text-[10px] px-2 py-0.5 bg-gray-200 text-gray-600 rounded hover:bg-gray-300"
                        >
                          Cancel
                        </button>
                      </div>
                    </div>
                  ) : (
                    <button
                      onClick={() => { setAddingToHalfTerm(ht.id); setNewUnitTitle(''); }}
                      className="mt-1 flex items-center gap-0.5 text-[10px] text-gray-400 hover:text-blue-600 transition-colors"
                    >
                      <Plus size={10} /> Add unit
                    </button>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Comment edit modal */}
      {editingUnitId && (() => {
        const unit = data.units.find((u) => u.id === editingUnitId)!;
        return (
          <div
            className="fixed inset-0 bg-black/30 flex items-center justify-center z-50"
            onClick={() => setEditingUnitId(null)}
          >
            <div
              className="bg-white rounded-xl shadow-xl p-6 w-full max-w-md"
              onClick={(e) => e.stopPropagation()}
            >
              <h2 className="text-sm font-semibold text-gray-900 mb-1">{unit.title}</h2>
              <p className="text-xs text-gray-500 mb-3">Add a note or comment about this unit.</p>
              <textarea
                autoFocus
                rows={4}
                value={editComment}
                onChange={(e) => setEditComment(e.target.value)}
                className="w-full text-sm border border-gray-300 rounded-lg p-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500 resize-none"
                placeholder="Write a comment…"
              />
              <div className="flex justify-end gap-2 mt-3">
                <button
                  onClick={() => setEditingUnitId(null)}
                  className="px-3 py-1.5 text-sm text-gray-600 bg-gray-100 rounded-lg hover:bg-gray-200"
                >
                  Cancel
                </button>
                <button
                  onClick={() => handleSaveComment(editingUnitId)}
                  className="px-3 py-1.5 text-sm text-white bg-blue-600 rounded-lg hover:bg-blue-700"
                >
                  Save
                </button>
              </div>
            </div>
          </div>
        );
      })()}
    </div>
  );
}

function currentAcademicYear(): string {
  const now = new Date();
  const year = now.getFullYear();
  return now.getMonth() >= 8
    ? `${year}/${String(year + 1).slice(2)}`
    : `${year - 1}/${String(year).slice(2)}`;
}
