import React from 'react';
import { Plus, Layers, FileSpreadsheet } from 'lucide-react';
import { Sheet } from '../types';

interface TabsProps {
  sheets: { [id: string]: Sheet };
  sheetIds: string[];
  activeSheetId: string;
  onSelectSheet: (id: string) => void;
  onAddSheet: () => void;
}

export const Tabs: React.FC<TabsProps> = ({
  sheets,
  sheetIds,
  activeSheetId,
  onSelectSheet,
  onAddSheet,
}) => {
  return (
    <div className="flex items-center justify-between px-4 py-1.5 bg-slate-50 border-t border-slate-200 text-slate-700">
      <div className="flex items-center gap-1 overflow-x-auto scrollbar-none pr-4">
        {sheetIds.map((id) => {
          const sheet = sheets[id];
          const isActive = id === activeSheetId;
          return (
            <button
              key={id}
              onClick={() => onSelectSheet(id)}
              className={`flex items-center gap-1.5 px-3.5 py-1 text-xs font-semibold rounded-t-md transition-all border-x border-t cursor-pointer ${
                isActive
                  ? 'bg-white text-emerald-700 border-slate-200 shadow-sm border-b-2 border-b-emerald-600 font-bold'
                  : 'bg-slate-100/70 hover:bg-slate-100 text-slate-500 hover:text-slate-700 border-transparent border-b-slate-200'
              }`}
            >
              <FileSpreadsheet className={`w-3.5 h-3.5 ${isActive ? 'text-emerald-600' : 'text-slate-400'}`} />
              <span>{sheet?.name || id}</span>
            </button>
          );
        })}

        <button
          onClick={onAddSheet}
          className="p-1 hover:bg-slate-200/80 rounded transition-colors text-slate-400 hover:text-slate-700 cursor-pointer"
          title="Add a new sheet"
        >
          <Plus className="w-4 h-4" />
        </button>
      </div>

      {/* Summary active cells info indicator */}
      <div className="hidden sm:flex items-center gap-1.5 text-[10px] uppercase font-mono tracking-wider font-bold text-slate-400">
        <Layers className="w-3.5 h-3.5" />
        <span>Active Grid size: {Object.keys(sheets[activeSheetId]?.data || {}).length} populated cells</span>
      </div>
    </div>
  );
};
