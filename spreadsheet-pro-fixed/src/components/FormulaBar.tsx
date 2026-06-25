import React, { useState, useEffect } from 'react';

interface FormulaBarProps {
  selectedCellId: string | null;
  cellNameInput: string;
  onCellNameInputChange: (val: string) => void;
  onCellNameInputSubmit: () => void;
  value: string;
  onChange: (newValue: string) => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLInputElement>) => void;
}

export const FormulaBar: React.FC<FormulaBarProps> = ({
  selectedCellId,
  cellNameInput,
  onCellNameInputChange,
  onCellNameInputSubmit,
  value,
  onChange,
  onKeyDown
}) => {
  const [showSuggestions, setShowSuggestions] = useState(false);
  const FORMULAS = [
    "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "ABS", "SQRT", "ROUND", 
    "LEN", "UPPER", "LOWER", "CONCAT", "IF", "AND", "OR", "NOT", 
    "TODAY", "NOW", "VLOOKUP", "XLOOKUP", "SUMIF", "COUNTIF", 
    "AVERAGEIF", "ISBLANK", "ISNUMBER"
  ];

  const getSuggestions = (val: string) => {
    if (!val || (!val.startsWith("=") && !val.startsWith("+"))) return [];
    const match = val.match(/([A-Z_]+)$/i);
    if (!match) return [];
    const typed = match[1].toUpperCase();
    return FORMULAS.filter(f => f.startsWith(typed) && f !== typed);
  };

  const suggestions = getSuggestions(value);

  const selectSuggestion = (sug: string) => {
    const match = value.match(/([A-Z_]+)$/i);
    if (!match) return;
    const index = match.index ?? 0;
    const newVal = value.slice(0, index) + sug + "(";
    onChange(newVal);
    setShowSuggestions(false);
  };

  return (
    <div className="flex items-center gap-2 px-4 py-1.5 bg-slate-50 border-b border-slate-200">
      {/* Target Cell Label Input */}
      <input
        type="text"
        value={cellNameInput}
        onChange={(e) => onCellNameInputChange(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter') {
            onCellNameInputSubmit();
            e.currentTarget.blur();
          }
        }}
        placeholder="Cell"
        title="Type coordinate (e.g. B10) and press Enter to jump"
        className="w-[70px] h-7 bg-white border border-slate-250 rounded-md font-mono text-xs text-slate-700 font-bold text-center tracking-tight shadow-3xs focus:outline-hidden focus:ring-1 focus:ring-green-500"
      />

      {/* FX Icon Divider */}
      <span className="text-slate-400 font-serif font-semibold italic text-sm select-none px-1">
        ƒx
      </span>

      {/* Formula Entry Input Field */}
      <div className="flex-1 relative">
        <input
          type="text"
          id="formula-bar-input"
          value={value}
          onChange={(e) => {
            onChange(e.target.value);
            setShowSuggestions(true);
          }}
          onFocus={() => setShowSuggestions(true)}
          onBlur={() => setTimeout(() => setShowSuggestions(false), 250)}
          onKeyDown={onKeyDown}
          placeholder={selectedCellId ? "Enter numbers, text, or formulas starting with = or +" : "Select a cell to enter values or formulas"}
          disabled={!selectedCellId}
          className="w-full px-3 py-1 text-xs font-mono bg-white border border-slate-200 rounded-md shadow-3xs focus:outline-hidden focus:ring-1 focus:ring-green-500 text-slate-800 disabled:bg-slate-100 disabled:text-slate-400"
        />
        {showSuggestions && suggestions.length > 0 && (
          <div className="absolute left-0 right-0 top-full mt-1 z-50 bg-white border border-slate-250 rounded-md shadow-lg max-h-48 overflow-y-auto py-1">
            {suggestions.map(sug => (
              <button
                key={sug}
                type="button"
                onMouseDown={() => selectSuggestion(sug)}
                className="w-full text-left px-3 py-1.5 text-xs font-mono hover:bg-slate-100 text-slate-700 cursor-pointer flex items-center justify-between"
              >
                <span>{sug}</span>
                <span className="text-[10px] text-slate-400">Formula</span>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};
