'use client';

import { useState, useRef, useEffect, useCallback } from 'react';

const TEMPLATE_VARIABLES = [
  { label: '<Name>', description: 'Pupil first name' },
  { label: '<He>', description: 'He/She (capitalised)' },
  { label: '<he>', description: 'he/she (lowercase)' },
  { label: '<Him>', description: 'Him/Her (capitalised)' },
  { label: '<him>', description: 'him/her (lowercase)' },
  { label: '<His>', description: 'His/Her (capitalised)' },
  { label: '<his>', description: 'his/her (lowercase)' },
  { label: '<Subject>', description: 'Subject title' },
  { label: '<Year>', description: 'Year group' },
  { label: '<EoYLevel>', description: 'End of year level' },
  { label: '<TargetLevel>', description: 'Target level' },
];

interface VariableAutocompleteInputProps {
  name: string;
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  required?: boolean;
  className?: string;
  autoFocus?: boolean;
}

export default function VariableAutocompleteInput({
  name,
  value,
  onChange,
  placeholder,
  required,
  className = '',
  autoFocus,
}: VariableAutocompleteInputProps) {
  const [showDropdown, setShowDropdown] = useState(false);
  const [filter, setFilter] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(0);
  const [triggerPos, setTriggerPos] = useState<number | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const dropdownRef = useRef<HTMLDivElement>(null);

  const filtered = TEMPLATE_VARIABLES.filter(v =>
    filter === '' || v.label.toLowerCase().includes(filter.toLowerCase())
  );

  // Close dropdown on outside click
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(e.target as Node) &&
        inputRef.current &&
        !inputRef.current.contains(e.target as Node)
      ) {
        setShowDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const insertVariable = useCallback((variable: string) => {
    if (triggerPos === null) return;

    // Replace from the `<` trigger position to current cursor, with the full variable
    const before = value.substring(0, triggerPos);
    const cursorPos = inputRef.current?.selectionStart ?? value.length;
    const after = value.substring(cursorPos);

    const newValue = before + variable + after;
    onChange(newValue);
    setShowDropdown(false);
    setTriggerPos(null);

    // Restore cursor position after React re-render
    requestAnimationFrame(() => {
      if (inputRef.current) {
        const newCursorPos = before.length + variable.length;
        inputRef.current.setSelectionRange(newCursorPos, newCursorPos);
        inputRef.current.focus();
      }
    });
  }, [triggerPos, value, onChange]);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const newValue = e.target.value;
    const cursorPos = e.target.selectionStart ?? newValue.length;
    onChange(newValue);

    // Check if we should open the dropdown
    // Find the last unmatched `<` before the cursor
    const textBeforeCursor = newValue.substring(0, cursorPos);
    const lastAngle = textBeforeCursor.lastIndexOf('<');

    if (lastAngle !== -1) {
      const textAfterAngle = textBeforeCursor.substring(lastAngle + 1);
      // Only show if there's no `>` between the `<` and cursor (i.e. not a closed tag)
      if (!textAfterAngle.includes('>')) {
        setTriggerPos(lastAngle);
        setFilter('<' + textAfterAngle);
        setShowDropdown(true);
        setSelectedIndex(0);
        return;
      }
    }

    setShowDropdown(false);
    setTriggerPos(null);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!showDropdown || filtered.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setSelectedIndex(i => (i + 1) % filtered.length);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setSelectedIndex(i => (i - 1 + filtered.length) % filtered.length);
    } else if (e.key === 'Enter' || e.key === 'Tab') {
      if (showDropdown && filtered.length > 0) {
        e.preventDefault();
        insertVariable(filtered[selectedIndex].label);
      }
    } else if (e.key === 'Escape') {
      setShowDropdown(false);
    }
  };

  return (
    <div className="relative flex-1">
      {/* Hidden input to submit the value as form data */}
      <input type="hidden" name={name} value={value} />
      <input
        ref={inputRef}
        type="text"
        value={value}
        onChange={handleChange}
        onKeyDown={handleKeyDown}
        placeholder={placeholder}
        required={required}
        autoFocus={autoFocus}
        className={className || 'w-full px-2 py-1.5 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-primary focus:border-transparent'}
      />
      {showDropdown && filtered.length > 0 && (
        <div
          ref={dropdownRef}
          className="absolute left-0 top-full mt-1 w-72 bg-white border border-gray-200 rounded-lg shadow-lg z-50 max-h-64 overflow-y-auto"
        >
          {filtered.map((variable, index) => (
            <button
              key={variable.label}
              type="button"
              onMouseDown={(e) => {
                e.preventDefault(); // Prevent blur
                insertVariable(variable.label);
              }}
              onMouseEnter={() => setSelectedIndex(index)}
              className={`w-full text-left px-3 py-2 flex items-center gap-3 text-sm transition-colors ${
                index === selectedIndex
                  ? 'bg-primary/10 text-primary'
                  : 'hover:bg-gray-50 text-gray-700'
              }`}
            >
              <code className="text-xs font-mono bg-gray-100 px-1.5 py-0.5 rounded border border-gray-200 whitespace-nowrap">
                {variable.label}
              </code>
              <span className="text-xs text-gray-500 truncate">{variable.description}</span>
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
