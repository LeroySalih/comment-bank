'use client';

import { useState, useRef } from 'react';
import { uploadClassLinkedData } from '@/lib/server-actions/linked-data';

interface LinkedDataUploadProps {
  classes: { id: string; name: string }[];
}

export function LinkedDataUpload({ classes }: LinkedDataUploadProps) {
  const [selectedClassId, setSelectedClassId] = useState('');
  const [uploading, setUploading] = useState(false);
  const [result, setResult] = useState<{
    matchedCount: number;
    unmatchedRows: string[];
    detectedFields: string[];
  } | null>(null);
  const [error, setError] = useState('');
  const fileRef = useRef<HTMLInputElement>(null);

  const handleUpload = async () => {
    if (!selectedClassId) {
      setError('Please select a class');
      return;
    }

    const file = fileRef.current?.files?.[0];
    if (!file) {
      setError('Please select a file');
      return;
    }

    setUploading(true);
    setError('');
    setResult(null);

    try {
      const buffer = await file.arrayBuffer();
      const base64 = btoa(
        new Uint8Array(buffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
      );

      const response = await uploadClassLinkedData(selectedClassId, base64);

      if (response.success) {
        setResult({
          matchedCount: (response as any).matchedCount,
          unmatchedRows: (response as any).unmatchedRows || [],
          detectedFields: (response as any).detectedFields || []
        });
        if (fileRef.current) fileRef.current.value = '';
      } else {
        setError((response as any).error || 'Upload failed');
      }
    } catch (err) {
      setError('An unexpected error occurred');
    }

    setUploading(false);
  };

  return (
    <div className="space-y-6">
      <div>
        <h2 className="text-xl font-bold text-gray-900">Upload Linked Data</h2>
        <p className="text-sm text-gray-500 mt-1">
          Upload an Excel or CSV file to populate linked data for a class. The file should have an AdmissionNumber column and any data columns (e.g. Behaviour, Effort, Homework).
        </p>
      </div>

      <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Class</label>
          <select
            value={selectedClassId}
            onChange={(e) => setSelectedClassId(e.target.value)}
            className="w-full max-w-sm px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-primary focus:border-transparent"
          >
            <option value="">Select a class...</option>
            {classes.map(c => (
              <option key={c.id} value={c.id}>{c.name}</option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">File (Excel or CSV)</label>
          <input
            ref={fileRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="block w-full max-w-sm text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-medium file:bg-primary/10 file:text-primary hover:file:bg-primary/20"
          />
        </div>

        <button
          onClick={handleUpload}
          disabled={uploading}
          className="px-4 py-2 text-sm font-medium text-white bg-primary rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
        >
          {uploading ? 'Uploading...' : 'Upload'}
        </button>
      </div>

      {error && (
        <div className="p-4 bg-red-50 border border-red-200 rounded-lg text-sm text-red-700">
          {error}
        </div>
      )}

      {result && (
        <div className="bg-white border border-gray-200 rounded-xl p-6 shadow-sm space-y-4">
          <h3 className="text-base font-bold text-gray-900">Upload Result</h3>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="bg-green-50 border border-green-200 rounded-lg p-4 text-center">
              <p className="text-2xl font-bold text-green-700">{result.matchedCount}</p>
              <p className="text-sm text-green-600">Assignments Updated</p>
            </div>
            <div className="bg-amber-50 border border-amber-200 rounded-lg p-4 text-center">
              <p className="text-2xl font-bold text-amber-700">{result.unmatchedRows.length}</p>
              <p className="text-sm text-amber-600">Unmatched Rows</p>
            </div>
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4 text-center">
              <p className="text-2xl font-bold text-blue-700">{result.detectedFields.length}</p>
              <p className="text-sm text-blue-600">Fields Detected</p>
            </div>
          </div>

          {result.detectedFields.length > 0 && (
            <div>
              <p className="text-sm font-medium text-gray-700 mb-2">Detected Fields:</p>
              <div className="flex flex-wrap gap-2">
                {result.detectedFields.map(f => (
                  <span key={f} className="inline-flex items-center gap-1 text-xs font-medium text-purple-700 bg-purple-100 px-2 py-1 rounded">
                    <span className="material-symbols-outlined text-sm">link</span>
                    {f}
                  </span>
                ))}
              </div>
            </div>
          )}

          {result.unmatchedRows.length > 0 && (
            <div>
              <p className="text-sm font-medium text-gray-700 mb-2">Unmatched Admission Numbers:</p>
              <div className="max-h-40 overflow-y-auto bg-gray-50 rounded-lg p-3">
                <ul className="text-xs text-gray-600 space-y-1">
                  {result.unmatchedRows.slice(0, 50).map((r, i) => (
                    <li key={i}>{r}</li>
                  ))}
                  {result.unmatchedRows.length > 50 && (
                    <li className="font-medium">...and {result.unmatchedRows.length - 50} more</li>
                  )}
                </ul>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
