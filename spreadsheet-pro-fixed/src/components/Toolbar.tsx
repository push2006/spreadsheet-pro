import React, { useState } from 'react';
import { 
  Bold, 
  Italic, 
  Underline, 
  AlignLeft, 
  AlignCenter, 
  AlignRight, 
  Undo2, 
  Redo2, 
  Download, 
  Plus, 
  Trash2, 
  Edit3,
  Search,
  ArrowUpDown,
  Zap,
  Info,
  Upload,
  Grid3X3,
  Palette,
  Brush,
  Sparkles,
  Calendar,
  Play,
  FileText,
  StickyNote,
  BarChart2,
  Minimize2,
  Trash,
  Copy,
  Scissors,
  Clipboard,
  Layers,
  Database,
  Columns4,
  CopyCheck,
  Share2,
  FileCode
} from 'lucide-react';
import { CellStyle, SelectionRange } from '../types';

interface ToolbarProps {
  activeSheetName: string;
  onAddSheet: () => void;
  onRenameSheet: (name: string) => void;
  onDeleteSheet: () => void;
  canUndo: boolean;
  canRedo: boolean;
  onUndo: () => void;
  onRedo: () => void;
  onExportCSV: () => void;
  selectedCellStyle: CellStyle;
  onUpdateStyle: (style: Partial<CellStyle>) => void;
  onSortActiveCol: (ascending: boolean) => void;
  onFindReplace: (findText: string, replaceText: string) => void;
  selectedCellId: string | null;
  onAutoFillValue: () => void; // Replicate starting value across selection
  hasRangeSelected: boolean;
  onOpenImport: () => void;
  onInsertRow: (row: number) => void;
  onDeleteRow: (row: number) => void;
  onInsertColumn: (col: string) => void;
  onDeleteColumn: (col: string) => void;

  // New Workbook customize themes & layout & automate props
  themeColor: 'green' | 'blue' | 'purple' | 'rose' | 'amber' | 'slate';
  onThemeColorChange: (color: 'green' | 'blue' | 'purple' | 'rose' | 'amber' | 'slate') => void;
  hideGridlines: boolean;
  onHideGridlinesChange: (val: boolean) => void;
  drawModeEnabled: boolean;
  onDrawModeEnabledChange: (val: boolean) => void;
  drawColor: string;
  onDrawColorChange: (color: string) => void;
  drawThickness: number;
  onDrawThicknessChange: (thickness: number) => void;
  onClearDrawings: () => void;
  onInsertFloatingWidget: (type: 'chart' | 'note' | 'shape' | 'slicer' | 'pivot' | 'wordart', extra?: any) => void;
  activeRange: SelectionRange | null;
  onRunMacroScript: (scriptText: string) => void;
  onInsertFormulaValue: (formulaTemplate: string) => void;
  showFormulasActive: boolean;
  onShowFormulasActiveChange: (val: boolean) => void;

  // Clipboard & Format Painter features
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onPasteSpecial: (mode: 'all' | 'values' | 'formats' | 'formulas' | 'transpose') => void;
  hasClipboard: boolean;
  isFormatPainterActive: boolean;
  onToggleFormatPainter: () => void;
  onDuplicateSheet: () => void;

  // New Workbook Backup & Data Tools features
  onCopyWorkbookBackup: () => void;
  onPasteWorkbookBackup: (jsonString: string) => void;
  onRemoveDuplicates: () => void;
  onTextToColumns: (delimiter: string) => void;
  onCopySheetCSVText: () => void;
}

export const Toolbar: React.FC<ToolbarProps> = ({
  activeSheetName,
  onAddSheet,
  onRenameSheet,
  onDeleteSheet,
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  onExportCSV,
  selectedCellStyle,
  onUpdateStyle,
  onSortActiveCol,
  onFindReplace,
  selectedCellId,
  onAutoFillValue,
  hasRangeSelected,
  onOpenImport,
  onInsertRow,
  onDeleteRow,
  onInsertColumn,
  onDeleteColumn,

  themeColor,
  onThemeColorChange,
  hideGridlines,
  onHideGridlinesChange,
  drawModeEnabled,
  onDrawModeEnabledChange,
  drawColor,
  onDrawColorChange,
  drawThickness,
  onDrawThicknessChange,
  onClearDrawings,
  onInsertFloatingWidget,
  activeRange,
  onRunMacroScript,
  onInsertFormulaValue,
  showFormulasActive,
  onShowFormulasActiveChange,

  // Clipboard & Format Painter destructured
  onCopy,
  onCut,
  onPaste,
  onPasteSpecial,
  hasClipboard,
  isFormatPainterActive,
  onToggleFormatPainter,
  onDuplicateSheet,

  // New Backup & Data Tools destructured
  onCopyWorkbookBackup,
  onPasteWorkbookBackup,
  onRemoveDuplicates,
  onTextToColumns,
  onCopySheetCSVText
}) => {
  const [isRenaming, setIsRenaming] = useState(false);
  const [newName, setNewName] = useState(activeSheetName);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [showFindReplace, setShowFindReplace] = useState(false);

  // Active Ribbon Tab (Excel-style)
  const [activeRibbonTab, setActiveRibbonTab] = useState<'home' | 'insert' | 'layout' | 'draw' | 'formulas' | 'automate'>('home');

  // Custom picker helpers
  const parsedActive = selectedCellId ? selectedCellId.toUpperCase().match(/^([A-Z]+)(\d+)$/) : null;
  const activeColName = parsedActive ? parsedActive[1] : null;
  const activeRowIndex = parsedActive ? parseInt(parsedActive[2], 10) : null;

  // Macro templates and state
  const [macroScript, setMacroScript] = useState<string>(
    `// Macro: Multiply numbers in range A1:A5 by 2\nfor (let r = 1; r <= 5; r++) {\n  let id = "A" + r;\n  let val = parseFloat(getCell(id) || 0);\n  if (!isNaN(val)) {\n    setCell(id, String(val * 2));\n  }\n}`
  );

  // Custom date picker target state
  const [calendarDate, setCalendarDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [splitDelimiter, setSplitDelimiter] = useState<string>(",");
  const [backupPasteText, setBackupPasteText] = useState<string>("");
  const [showRestorePanel, setShowRestorePanel] = useState<boolean>(false);

  // States for advanced insert tools
  const [slicerCol, setSlicerCol] = useState('A');
  const [pivotRowCol, setPivotRowCol] = useState('A');
  const [pivotValCol, setPivotValCol] = useState('B');
  const [pivotOp, setPivotOp] = useState<'sum' | 'avg' | 'count' | 'max' | 'min'>('sum');
  const [wordartText, setWordartText] = useState('EXCEL PRO');
  const [wordartColor, setWordartColor] = useState('#6366f1');

  const handleRenameSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (newName.trim()) {
      onRenameSheet(newName.trim());
      setIsRenaming(false);
    }
  };

  // Helper theme styles
  const getThemeColorClass = () => {
    switch (themeColor) {
      case 'blue': return 'bg-blue-600 hover:bg-blue-700';
      case 'purple': return 'bg-purple-600 hover:bg-purple-700';
      case 'rose': return 'bg-rose-600 hover:bg-rose-700';
      case 'amber': return 'bg-amber-600 hover:bg-amber-700';
      case 'slate': return 'bg-slate-700 hover:bg-slate-800';
      case 'green':
      default:
        return 'bg-green-600 hover:bg-green-700';
    }
  };

  const getThemeTextClass = () => {
    switch (themeColor) {
      case 'blue': return 'text-blue-600';
      case 'purple': return 'text-purple-600';
      case 'rose': return 'text-rose-600';
      case 'amber': return 'text-amber-600';
      case 'slate': return 'text-slate-700';
      case 'green':
      default:
        return 'text-green-600';
    }
  };

  const getThemeBorderClass = () => {
    switch (themeColor) {
      case 'blue': return 'border-blue-500';
      case 'purple': return 'border-purple-500';
      case 'rose': return 'border-rose-500';
      case 'amber': return 'border-amber-500';
      case 'slate': return 'border-slate-500';
      case 'green':
      default:
        return 'border-green-500';
    }
  };

  const getThemeBgClass = () => {
    switch (themeColor) {
      case 'blue': return 'bg-blue-50';
      case 'purple': return 'bg-purple-50';
      case 'rose': return 'bg-rose-50';
      case 'amber': return 'bg-amber-50';
      case 'slate': return 'bg-slate-50';
      case 'green':
      default:
        return 'bg-green-50';
    }
  };

  const getActiveRangeString = () => {
    if (!activeRange) return selectedCellId || "";
    return `${activeRange.start}:${activeRange.end}`;
  };

  return (
    <div className="flex flex-col border-b border-slate-200 bg-white shadow-xs">
      {/* UPPER BAR: File Name / Sheet Operations */}
      <div className="flex flex-wrap items-center justify-between gap-2 px-4 py-1.5 border-b border-slate-100">
        <div className="flex items-center gap-3">
          {/* Active Sheet Selector Name */}
          <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-md px-2 py-0.5">
            {isRenaming ? (
              <form onSubmit={handleRenameSubmit} className="flex items-center gap-1">
                <input 
                  type="text" 
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                  className={`px-1.5 py-0.5 text-xs bg-white border border-slate-300 rounded focus:outline-hidden focus:ring-1 focus:ring-slate-400 max-w-[100px]`}
                  autoFocus
                />
                <button type="submit" className="text-[10px] text-slate-700 hover:bg-slate-150 font-semibold px-1 py-0.5 border border-slate-200 bg-white rounded">Save</button>
                <button type="button" onClick={() => setIsRenaming(false)} className="text-[10px] text-slate-500 hover:text-slate-600 px-1 py-0.5 border border-slate-200 bg-white rounded">Cancel</button>
              </form>
            ) : (
              <>
                <span className="text-xs font-semibold text-slate-700">Active Sheet: <span className={`${getThemeTextClass()} underline decoration-dotted font-bold`}>{activeSheetName}</span></span>
                <button 
                  onClick={() => { setNewName(activeSheetName); setIsRenaming(true); }}
                  className="p-0.5 hover:bg-slate-200 text-slate-500 hover:text-slate-800 rounded transition-colors"
                  title="Rename active sheet"
                >
                  <Edit3 className="w-3 h-3" />
                </button>
              </>
            )}
          </div>

          <button 
            onClick={onAddSheet}
            className={`flex items-center gap-1 px-2 py-1 text-xs font-medium text-white ${getThemeColorClass()} rounded-md shadow-xs transition-colors cursor-pointer`}
            title="Create a new sheet"
          >
            <Plus className="w-3.5 h-3.5" />
            <span>Add Sheet</span>
          </button>

          <button 
            onClick={onDuplicateSheet}
            className="flex items-center gap-1 px-2 py-1 text-xs font-medium text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-md shadow-3xs transition-colors cursor-pointer"
            title="Duplicate (make a copy of) active sheet"
          >
            <Layers className="w-3.5 h-3.5 text-indigo-500" />
            <span>Duplicate</span>
          </button>

          <button 
            onClick={onDeleteSheet}
            className="p-1 hover:bg-red-50 text-slate-400 hover:text-red-600 border border-transparent hover:border-red-150 rounded-md transition-all cursor-pointer"
            title="Delete this sheet"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        </div>

        {/* Global actions: Documents / Import & Export / Search */}
        <div className="flex items-center gap-2">
          {/* Import */}
          <button
            onClick={onOpenImport}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-md shadow-3xs transition-colors cursor-pointer"
            title="Import TXT, CSV, PDF, or Word Document patterns"
          >
            <Upload className="w-3.5 h-3.5 text-slate-500" />
            <span>Import</span>
          </button>

          {/* Find & Replace Toggle */}
          <button
            onClick={() => setShowFindReplace(!showFindReplace)}
            className={`flex items-center gap-1 px-2.5 py-1 text-xs font-medium rounded-md border transition-colors cursor-pointer ${
              showFindReplace 
                ? `${getThemeBgClass()} ${getThemeTextClass()} ${getThemeBorderClass()}` 
                : 'bg-white hover:bg-slate-50 text-slate-700 border-slate-200'
            }`}
            title="Find & Replace values in active spreadsheet"
          >
            <Search className="w-3.5 h-3.5" />
            <span>Find & Replace</span>
          </button>

          {/* Export */}
          <button 
            onClick={onExportCSV}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-slate-700 bg-white hover:bg-slate-50 border border-slate-300 rounded-md shadow-2xs transition-colors cursor-pointer"
            title="Export this sheet as CSV file"
          >
            <Download className="w-3.5 h-3.5 text-slate-500" />
            <span>Export CSV</span>
          </button>
        </div>
      </div>

      {/* RIBBON TABS SELECTION MENU */}
      <div className="flex items-center bg-slate-100/80 px-4 border-b border-slate-200/80">
        <button
          onClick={() => setActiveRibbonTab('home')}
          className={`px-3 py-1.5 text-xs font-bold transition-all border-b-2 cursor-pointer ${
            activeRibbonTab === 'home' 
              ? `${getThemeBorderClass()} ${getThemeTextClass()} bg-white border-t border-x border-slate-200 rounded-t-md -mb-[1px]` 
              : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
        >
          Home
        </button>
        <button
          onClick={() => setActiveRibbonTab('insert')}
          className={`px-3 py-1.5 text-xs font-bold transition-all border-b-2 cursor-pointer ${
            activeRibbonTab === 'insert' 
              ? `${getThemeBorderClass()} ${getThemeTextClass()} bg-white border-t border-x border-slate-200 rounded-t-md -mb-[1px]` 
              : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
        >
          Insert
        </button>
        <button
          onClick={() => setActiveRibbonTab('layout')}
          className={`px-3 py-1.5 text-xs font-bold transition-all border-b-2 cursor-pointer ${
            activeRibbonTab === 'layout' 
              ? `${getThemeBorderClass()} ${getThemeTextClass()} bg-white border-t border-x border-slate-200 rounded-t-md -mb-[1px]` 
              : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
        >
          Page Layout
        </button>
        <button
          onClick={() => setActiveRibbonTab('draw')}
          className={`px-3 py-1.5 text-xs font-bold transition-all border-b-2 cursor-pointer ${
            activeRibbonTab === 'draw' 
              ? `${getThemeBorderClass()} ${getThemeTextClass()} bg-white border-t border-x border-slate-200 rounded-t-md -mb-[1px]` 
              : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
        >
          Draw
        </button>
        <button
          onClick={() => setActiveRibbonTab('formulas')}
          className={`px-3 py-1.5 text-xs font-bold transition-all border-b-2 cursor-pointer ${
            activeRibbonTab === 'formulas' 
              ? `${getThemeBorderClass()} ${getThemeTextClass()} bg-white border-t border-x border-slate-200 rounded-t-md -mb-[1px]` 
              : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
        >
          Formulas & Date
        </button>
        <button
          onClick={() => setActiveRibbonTab('automate')}
          className={`px-3 py-1.5 text-xs font-bold transition-all border-b-2 cursor-pointer ${
            activeRibbonTab === 'automate' 
              ? `${getThemeBorderClass()} ${getThemeTextClass()} bg-white border-t border-x border-slate-200 rounded-t-md -mb-[1px]` 
              : 'border-transparent text-slate-600 hover:text-slate-900 hover:bg-slate-200/50'
          }`}
        >
          Automate
        </button>
      </div>

      {/* DYNAMIC RIBBON CONTENTS */}
      <div className="flex items-center justify-between gap-x-4 px-4 py-2 bg-slate-50/50 min-h-[52px]">
        
        {/* TAB 1: HOME */}
        {activeRibbonTab === 'home' && (
          <div className="flex items-center flex-wrap gap-x-4 gap-y-1.5 w-full">
            
            {/* Group: Undo Redo */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-1">
              <button 
                onClick={onUndo}
                disabled={!canUndo}
                className={`p-1 rounded hover:bg-slate-200 transition-colors cursor-pointer ${!canUndo ? 'opacity-30 cursor-not-allowed' : 'text-slate-700'}`}
                title="Undo (Ctrl+Z)"
              >
                <Undo2 className="w-4 h-4" />
              </button>
              <button 
                onClick={onRedo}
                disabled={!canRedo}
                className={`p-1 rounded hover:bg-slate-200 transition-colors cursor-pointer ${!canRedo ? 'opacity-30 cursor-not-allowed' : 'text-slate-700'}`}
                title="Redo (Ctrl+Y)"
              >
                <Redo2 className="w-4 h-4" />
              </button>
            </div>

            {/* Group: Clipboard */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-0.5">
              <button 
                onClick={onCopy}
                className="p-1 rounded hover:bg-slate-200 text-slate-500 hover:text-slate-800 transition-all cursor-pointer flex items-center gap-1"
                title="Copy (Ctrl+C)"
              >
                <Copy className="w-4 h-4 text-slate-600" />
              </button>
              <button 
                onClick={onCut}
                className="p-1 rounded hover:bg-slate-200 text-slate-500 hover:text-slate-800 transition-all cursor-pointer flex items-center gap-1"
                title="Cut (Ctrl+X)"
              >
                <Scissors className="w-4 h-4 text-slate-600" />
              </button>
              <button 
                onClick={onPaste}
                disabled={!hasClipboard}
                className={`p-1 rounded hover:bg-slate-200 transition-all cursor-pointer flex items-center gap-1 ${!hasClipboard ? 'opacity-30 cursor-not-allowed' : 'text-slate-700 hover:text-slate-800'}`}
                title="Paste (Ctrl+V)"
              >
                <Clipboard className="w-4 h-4 text-slate-600" />
              </button>
              
              {/* Paste Special Dropdown */}
              <div className="relative group/paste">
                <button 
                  disabled={!hasClipboard}
                  className={`px-1.5 py-0.5 text-[10px] rounded hover:bg-slate-200 border border-slate-200 transition-all cursor-pointer font-bold ${!hasClipboard ? 'opacity-30 cursor-not-allowed text-slate-400' : 'text-slate-600 hover:text-slate-800'}`}
                  title="Paste Special Options"
                >
                  Paste ▾
                </button>
                {hasClipboard && (
                  <div className="absolute top-full left-0 mt-1 bg-white border border-slate-200 shadow-lg rounded py-1 z-50 hidden group-hover/paste:block w-40">
                    <button
                      onClick={() => onPasteSpecial?.('all')}
                      className="w-full text-left px-3 py-1.5 text-xs hover:bg-slate-50 cursor-pointer text-slate-700 font-medium"
                    >
                      Paste All
                    </button>
                    <button
                      onClick={() => onPasteSpecial?.('values')}
                      className="w-full text-left px-3 py-1.5 text-xs hover:bg-slate-50 cursor-pointer text-slate-700 font-medium border-t border-slate-100"
                    >
                      Paste Values Only
                    </button>
                    <button
                      onClick={() => onPasteSpecial?.('formats')}
                      className="w-full text-left px-3 py-1.5 text-xs hover:bg-slate-50 cursor-pointer text-slate-700 font-medium"
                    >
                      Paste Formats Only
                    </button>
                    <button
                      onClick={() => onPasteSpecial?.('formulas')}
                      className="w-full text-left px-3 py-1.5 text-xs hover:bg-slate-50 cursor-pointer text-slate-700 font-medium border-t border-slate-100"
                    >
                      Paste Formulas Only
                    </button>
                    <button
                      onClick={() => onPasteSpecial?.('transpose')}
                      className="w-full text-left px-3 py-1.5 text-xs hover:bg-slate-50 cursor-pointer text-slate-700 font-medium text-indigo-600 font-semibold"
                    >
                      Paste Transposed ⇄
                    </button>
                  </div>
                )}
              </div>

              {/* Format Painter */}
              <button 
                onClick={onToggleFormatPainter}
                className={`p-1 rounded transition-all cursor-pointer ml-1 ${isFormatPainterActive ? 'bg-amber-100 text-amber-800 border border-amber-300 animate-pulse' : 'text-slate-500 hover:bg-slate-200/60'}`}
                title="Format Painter (Copy style & paint onto other cells)"
              >
                <Brush className="w-4 h-4" />
              </button>
            </div>

            {/* Group: Font Style */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-0.5">
              <button 
                onClick={() => onUpdateStyle({ bold: !selectedCellStyle.bold })}
                className={`p-1 rounded transition-all cursor-pointer ${selectedCellStyle.bold ? 'bg-slate-200 text-black font-bold' : 'text-slate-500 hover:bg-slate-200/60'}`}
                title="Bold"
              >
                <Bold className="w-4 h-4" />
              </button>
              <button 
                onClick={() => onUpdateStyle({ italic: !selectedCellStyle.italic })}
                className={`p-1 rounded transition-all cursor-pointer ${selectedCellStyle.italic ? 'bg-slate-200 text-black' : 'text-slate-500 hover:bg-slate-200/60'}`}
                title="Italic"
              >
                <Italic className="w-4 h-4" />
              </button>
              <button 
                onClick={() => onUpdateStyle({ underline: !selectedCellStyle.underline })}
                className={`p-1 rounded transition-all cursor-pointer ${selectedCellStyle.underline ? 'bg-slate-200 text-black' : 'text-slate-500 hover:bg-slate-200/60'}`}
                title="Underline"
              >
                <Underline className="w-4 h-4" />
              </button>
            </div>

            {/* Group: Alignments */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-0.5">
              <button 
                onClick={() => onUpdateStyle({ align: 'left' })}
                className={`p-1 rounded transition-colors cursor-pointer ${selectedCellStyle.align === 'left' || !selectedCellStyle.align ? 'bg-slate-200 text-slate-800' : 'text-slate-500 hover:bg-slate-200/60'}`}
                title="Align Left"
              >
                <AlignLeft className="w-4 h-4" />
              </button>
              <button 
                onClick={() => onUpdateStyle({ align: 'center' })}
                className={`p-1 rounded transition-colors cursor-pointer ${selectedCellStyle.align === 'center' ? 'bg-slate-200 text-slate-800' : 'text-slate-500 hover:bg-slate-200/60'}`}
                title="Align Center"
              >
                <AlignCenter className="w-4 h-4" />
              </button>
              <button 
                onClick={() => onUpdateStyle({ align: 'right' })}
                className={`p-1 rounded transition-colors cursor-pointer ${selectedCellStyle.align === 'right' ? 'bg-slate-200 text-slate-800' : 'text-slate-500 hover:bg-slate-200/60'}`}
                title="Align Right"
              >
                <AlignRight className="w-4 h-4" />
              </button>
            </div>

            {/* Group: Font Size */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-1">
              <span className="text-[9px] uppercase font-bold text-slate-400">Size</span>
              <select
                value={selectedCellStyle.fontSize || 13}
                onChange={(e) => onUpdateStyle({ fontSize: parseInt(e.target.value, 10) })}
                className="text-xs bg-white border border-slate-300 rounded px-1 py-0.5 focus:outline-hidden cursor-pointer"
              >
                {[10, 11, 12, 13, 14, 16, 18, 20, 24].map(sz => (
                  <option key={sz} value={sz}>{sz}px</option>
                ))}
              </select>
            </div>

            {/* Group: Custom Cell Color & Fill */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-2">
              <div className="flex items-center gap-1">
                <span className="text-[9px] uppercase font-bold text-slate-400">Fill</span>
                <input 
                  type="color" 
                  value={selectedCellStyle.bg || "#ffffff"}
                  onChange={(e) => onUpdateStyle({ bg: e.target.value })}
                  className="w-4.5 h-4.5 border border-slate-300 rounded-sm cursor-pointer p-0"
                  title="Background Fill Color"
                />
              </div>
              <div className="flex items-center gap-1">
                <span className="text-[9px] uppercase font-bold text-slate-400">Text</span>
                <input 
                  type="color" 
                  value={selectedCellStyle.color || "#0f172a"}
                  onChange={(e) => onUpdateStyle({ color: e.target.value })}
                  className="w-4.5 h-4.5 border border-slate-300 rounded-sm cursor-pointer p-0"
                  title="Text Color"
                />
              </div>
            </div>

            {/* Group: Format Selector */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-1">
              <span className="text-[9px] uppercase font-bold text-slate-400">Format</span>
              <select
                value={selectedCellStyle.format || 'general'}
                onChange={(e) => onUpdateStyle({ format: e.target.value as any })}
                className="text-xs bg-white border border-slate-300 rounded px-1.5 py-0.5 focus:outline-hidden cursor-pointer font-medium"
              >
                <option value="general">General</option>
                <option value="number">Number (1,234.56)</option>
                <option value="currency">Currency ($1,234.56)</option>
                <option value="percent">Percentage (12.3%)</option>
                <option value="scientific">Scientific (1.2e+3)</option>
              </select>
            </div>

            {/* Group: Borders Selector */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-1">
              <span className="text-[9px] uppercase font-bold text-slate-400">Border</span>
              <select
                value={selectedCellStyle.border || 'none'}
                onChange={(e) => onUpdateStyle({ border: e.target.value as any })}
                className="text-xs bg-white border border-slate-300 rounded px-1.5 py-0.5 focus:outline-hidden cursor-pointer"
              >
                <option value="none">None</option>
                <option value="all">All Borders</option>
                <option value="bottom">Bottom Border</option>
                <option value="top">Top Border</option>
                <option value="thick">Thick Border</option>
                <option value="double-bottom">Double Bottom</option>
              </select>
            </div>

            {/* Group: Edit Cells Actions */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-1">
              <span className="text-[9px] uppercase font-bold text-slate-400">Rows & Cols</span>
              <div className="flex gap-1">
                <button
                  type="button"
                  onClick={() => activeRowIndex !== null && onInsertRow(activeRowIndex)}
                  disabled={activeRowIndex === null}
                  className="px-1.5 py-0.5 text-[11px] bg-white hover:bg-slate-100 border border-slate-200 rounded disabled:opacity-40 text-green-700 font-bold"
                  title="Insert empty row above selected cell"
                >
                  + Row
                </button>
                <button
                  type="button"
                  onClick={() => activeRowIndex !== null && onDeleteRow(activeRowIndex)}
                  disabled={activeRowIndex === null}
                  className="px-1.5 py-0.5 text-[11px] bg-white hover:bg-red-50 border border-slate-200 rounded disabled:opacity-40 text-red-600"
                  title="Delete active row"
                >
                  - Row
                </button>
                <button
                  type="button"
                  onClick={() => activeColName !== null && onInsertColumn(activeColName)}
                  disabled={activeColName === null}
                  className="px-1.5 py-0.5 text-[11px] bg-white hover:bg-slate-100 border border-slate-200 rounded disabled:opacity-40 text-green-700 font-bold"
                  title="Insert empty column left of selected cell"
                >
                  + Col
                </button>
                <button
                  type="button"
                  onClick={() => activeColName !== null && onDeleteColumn(activeColName)}
                  disabled={activeColName === null}
                  className="px-1.5 py-0.5 text-[11px] bg-white hover:bg-red-50 border border-slate-200 rounded disabled:opacity-40 text-red-600"
                  title="Delete active column"
                >
                  - Col
                </button>
              </div>
            </div>

            {/* Group: Quick Styles */}
            <div className="flex items-center border-r border-slate-200 pr-3 gap-1">
              <span className="text-[9px] uppercase font-bold text-slate-400">Quick Format</span>
              <button
                type="button"
                onClick={() => onUpdateStyle({ format: 'currency' })}
                className="px-2 py-0.5 text-[11px] bg-white hover:bg-slate-100 border border-slate-200 rounded font-extrabold text-slate-700 cursor-pointer"
                title="Format selection as Currency ($)"
              >
                $
              </button>
              <button
                type="button"
                onClick={() => onUpdateStyle({ format: 'percent' })}
                className="px-2 py-0.5 text-[11px] bg-white hover:bg-slate-100 border border-slate-200 rounded font-extrabold text-slate-700 cursor-pointer"
                title="Format selection as Percentage (%)"
              >
                %
              </button>
              <button
                type="button"
                onClick={() => onUpdateStyle({ bold: false, italic: false, underline: false, bg: '#ffffff', color: '#0f172a', format: 'general', align: 'left', fontSize: 13, border: 'none' })}
                className="px-2 py-0.5 text-[11px] bg-white hover:bg-red-50 border border-red-200 rounded text-red-600 cursor-pointer font-medium"
                title="Clear formatting on selection"
              >
                Clear
              </button>
            </div>

            {/* Group: Sorting */}
            <div className="flex items-center gap-1.5">
              <span className="text-[9px] uppercase font-bold text-slate-400">Sort</span>
              <button 
                onClick={() => onSortActiveCol(true)}
                className="flex items-center gap-0.5 px-1.5 py-0.5 text-xs bg-white text-slate-700 hover:bg-slate-100 border border-slate-300 rounded cursor-pointer"
                title="Sort column alphabetically ascending"
              >
                <ArrowUpDown className="w-3 h-3 text-slate-400" />
                <span>A-Z</span>
              </button>
              <button 
                onClick={() => onSortActiveCol(false)}
                className="flex items-center gap-0.5 px-1.5 py-0.5 text-xs bg-white text-slate-700 hover:bg-slate-100 border border-slate-300 rounded cursor-pointer"
                title="Sort column alphabetically descending"
              >
                <ArrowUpDown className="w-3 h-3 text-slate-400 rotate-180" />
                <span>Z-A</span>
              </button>
            </div>

            {/* Selection operations: Autofill */}
            {hasRangeSelected && (
              <button
                onClick={onAutoFillValue}
                className={`flex items-center gap-1 ml-auto px-2 py-0.5 text-xs font-medium text-white ${getThemeColorClass()} rounded transition-colors cursor-pointer`}
                title="Replicate values from first selected cell onto entire range"
              >
                <Zap className="w-3 h-3" />
                <span>Auto-Fill (Ctrl+D)</span>
              </button>
            )}
          </div>
        )}

        {/* TAB 2: INSERT ELEMENTS */}
        {activeRibbonTab === 'insert' && (
          <div className="flex items-center flex-wrap gap-4 w-full text-slate-700">
            {/* Charts section */}
            <div className="flex items-center gap-2 border-r border-slate-200 pr-3">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <BarChart2 className="w-3.5 h-3.5 text-indigo-500" />
                Charts
              </span>
              <div className="flex gap-1">
                <button
                  onClick={() => onInsertFloatingWidget('chart', { chartType: 'bar', range: getActiveRangeString() })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-slate-700 cursor-pointer font-medium"
                >
                  Bar
                </button>
                <button
                  onClick={() => onInsertFloatingWidget('chart', { chartType: 'line', range: getActiveRangeString() })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-slate-700 cursor-pointer font-medium"
                >
                  Line
                </button>
                <button
                  onClick={() => onInsertFloatingWidget('chart', { chartType: 'area', range: getActiveRangeString() })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-slate-700 cursor-pointer font-medium"
                >
                  Area
                </button>
              </div>
            </div>

            {/* Notes Section */}
            <div className="flex items-center gap-2 border-r border-slate-200 pr-3">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <StickyNote className="w-3.5 h-3.5 text-amber-500" />
                Notes
              </span>
              <button
                onClick={() => onInsertFloatingWidget('note')}
                className="px-2 py-1 text-xs bg-amber-50 hover:bg-amber-100 border border-amber-200 text-amber-800 rounded font-medium shadow-3xs cursor-pointer"
              >
                + Sticky Note
              </button>
            </div>

            {/* Slicers Section */}
            <div className="flex items-center gap-2 border-r border-slate-200 pr-3">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <Grid3X3 className="w-3.5 h-3.5 text-emerald-500" />
                Slicer (Filter)
              </span>
              <div className="flex items-center gap-1">
                <select
                  value={slicerCol}
                  onChange={(e) => setSlicerCol(e.target.value)}
                  className="text-xs bg-white border border-slate-300 rounded px-1 py-0.5 focus:outline-hidden cursor-pointer"
                >
                  {Array.from({ length: 15 }, (_, i) => String.fromCharCode(65 + i)).map(c => (
                    <option key={c} value={c}>Col {c}</option>
                  ))}
                </select>
                <button
                  onClick={() => onInsertFloatingWidget('slicer', { range: slicerCol })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-emerald-700 font-bold cursor-pointer"
                >
                  + Slicer
                </button>
              </div>
            </div>

            {/* Pivot Table Section */}
            <div className="flex items-center gap-2 border-r border-slate-200 pr-3">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <FileText className="w-3.5 h-3.5 text-purple-500" />
                Pivot Table
              </span>
              <div className="flex items-center gap-1 text-[11px] text-slate-600 font-medium">
                <span>Row:</span>
                <select
                  value={pivotRowCol}
                  onChange={(e) => setPivotRowCol(e.target.value)}
                  className="text-[11px] bg-white border border-slate-300 rounded px-1 cursor-pointer"
                >
                  {Array.from({ length: 15 }, (_, i) => String.fromCharCode(65 + i)).map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
                <span>Val:</span>
                <select
                  value={pivotValCol}
                  onChange={(e) => setPivotValCol(e.target.value)}
                  className="text-[11px] bg-white border border-slate-300 rounded px-1 cursor-pointer"
                >
                  {Array.from({ length: 15 }, (_, i) => String.fromCharCode(65 + i)).map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
                <select
                  value={pivotOp}
                  onChange={(e) => setPivotOp(e.target.value as any)}
                  className="text-[11px] bg-white border border-slate-300 rounded px-1 cursor-pointer font-bold"
                >
                  <option value="sum">SUM</option>
                  <option value="avg">AVG</option>
                  <option value="count">COUNT</option>
                  <option value="max">MAX</option>
                  <option value="min">MIN</option>
                </select>
                <button
                  onClick={() => onInsertFloatingWidget('pivot', { 
                    range: getActiveRangeString() || 'A1:B10',
                    pivotRows: pivotRowCol,
                    pivotVals: pivotValCol,
                    pivotOp: pivotOp
                  })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-purple-700 font-bold cursor-pointer"
                  title="Insert a pivot table consolidating selection data groups"
                >
                  + Pivot
                </button>
              </div>
            </div>

            {/* Shapes Section */}
            <div className="flex items-center gap-2 border-r border-slate-200 pr-3">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <Palette className="w-3.5 h-3.5 text-rose-500" />
                Shapes
              </span>
              <div className="flex gap-1">
                <button
                  onClick={() => onInsertFloatingWidget('shape', { shapeType: 'rect', title: 'RECTANGLE' })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-slate-700 cursor-pointer font-medium"
                >
                  Rect
                </button>
                <button
                  onClick={() => onInsertFloatingWidget('shape', { shapeType: 'circle', title: 'CIRCLE' })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-slate-700 cursor-pointer font-medium"
                >
                  Circle
                </button>
                <button
                  onClick={() => onInsertFloatingWidget('shape', { shapeType: 'arrow', title: 'ARROW' })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-slate-700 cursor-pointer font-medium"
                >
                  Arrow
                </button>
              </div>
            </div>

            {/* WordArt Section */}
            <div className="flex items-center gap-2">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <Sparkles className="w-3.5 h-3.5 text-blue-500" />
                WordArt
              </span>
              <div className="flex items-center gap-1">
                <input
                  type="text"
                  value={wordartText}
                  onChange={(e) => setWordartText(e.target.value)}
                  className="text-xs bg-white border border-slate-300 rounded px-1.5 py-0.5 focus:outline-hidden w-20 font-bold"
                  placeholder="Text..."
                />
                <input
                  type="color"
                  value={wordartColor}
                  onChange={(e) => setWordartColor(e.target.value)}
                  className="w-4 h-4 cursor-pointer p-0 border border-slate-300 rounded-sm"
                />
                <button
                  onClick={() => onInsertFloatingWidget('wordart', { content: wordartText, color: wordartColor })}
                  className="px-2 py-1 text-xs bg-white hover:bg-slate-50 border border-slate-300 rounded shadow-3xs text-blue-700 font-bold cursor-pointer"
                >
                  + WordArt
                </button>
              </div>
            </div>
          </div>
        )}

        {/* TAB 3: PAGE LAYOUT */}
        {activeRibbonTab === 'layout' && (
          <div className="flex items-center flex-wrap gap-5 w-full">
            {/* Template Themes Customization */}
            <div className="flex items-center gap-2.5 border-r border-slate-200 pr-4">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <Palette className="w-4 h-4 text-slate-500" />
                Workbook Theme Template:
              </span>
              <div className="flex items-center gap-1.5 bg-white border border-slate-200 p-1 rounded-lg shadow-3xs">
                {[
                  { id: 'green', name: 'Emerald', hex: '#16a34a' },
                  { id: 'blue', name: 'Ocean Blue', hex: '#2563eb' },
                  { id: 'purple', name: 'Amethyst', hex: '#9333ea' },
                  { id: 'rose', name: 'Rose Petal', hex: '#e11d48' },
                  { id: 'amber', name: 'Gold Amber', hex: '#d97706' },
                  { id: 'slate', name: 'Slate Gray', hex: '#475569' }
                ].map((th) => (
                  <button
                    key={th.id}
                    onClick={() => onThemeColorChange(th.id as any)}
                    className={`w-5 h-5 rounded-full cursor-pointer relative transition-transform hover:scale-110 flex items-center justify-center ${
                      themeColor === th.id ? 'ring-2 ring-offset-1 ring-slate-400 scale-105' : ''
                    }`}
                    style={{ backgroundColor: th.hex }}
                    title={`Change template colour to ${th.name}`}
                  >
                    {themeColor === th.id && (
                      <span className="text-white text-[9px] font-bold">✓</span>
                    )}
                  </button>
                ))}
              </div>
            </div>

            {/* Gridlines toggler */}
            <div className="flex items-center gap-2 border-r border-slate-200 pr-4">
              <span className="text-[10px] uppercase font-bold text-slate-400">Sheet Gridlines</span>
              <button
                onClick={() => onHideGridlinesChange(!hideGridlines)}
                className={`px-3 py-1 text-xs font-semibold rounded border cursor-pointer transition-all flex items-center gap-1.5 ${
                  hideGridlines 
                    ? 'bg-slate-100 border-slate-300 text-slate-500' 
                    : `${getThemeBgClass()} ${getThemeBorderClass()} ${getThemeTextClass()}`
                }`}
              >
                <Grid3X3 className="w-3.5 h-3.5" />
                <span>{hideGridlines ? "Hidden (Hide)" : "Visible (Show)"}</span>
              </button>
            </div>

            {/* Page layout indicators */}
            <div className="flex items-center gap-4">
              <div className="flex items-center gap-1.5 text-xs text-slate-600">
                <span className="font-semibold text-slate-400 text-[10px] uppercase">Orientation:</span>
                <select className="bg-white border rounded px-1.5 py-0.5 text-xs focus:outline-hidden cursor-pointer font-medium text-slate-700">
                  <option value="portrait">Portrait Mode</option>
                  <option value="landscape">Landscape Mode</option>
                </select>
              </div>

              <div className="flex items-center gap-1.5 text-xs text-slate-600">
                <span className="font-semibold text-slate-400 text-[10px] uppercase">Margins:</span>
                <select className="bg-white border rounded px-1.5 py-0.5 text-xs focus:outline-hidden cursor-pointer font-medium text-slate-700">
                  <option value="normal">Normal (0.75")</option>
                  <option value="narrow">Narrow (0.25")</option>
                  <option value="wide">Wide (1.0")</option>
                </select>
              </div>
            </div>
          </div>
        )}

        {/* TAB 4: DRAW ON SHEET */}
        {activeRibbonTab === 'draw' && (
          <div className="flex items-center flex-wrap gap-5 w-full">
            <div className="flex items-center gap-3 border-r border-slate-200 pr-4">
              <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                <Brush className="w-4 h-4 text-slate-500" />
                Draw Pen Overlay
              </span>
              <button
                onClick={() => onDrawModeEnabledChange(!drawModeEnabled)}
                className={`px-3 py-1 text-xs font-bold rounded-lg cursor-pointer transition-colors ${
                  drawModeEnabled 
                    ? 'bg-rose-600 hover:bg-rose-700 text-white shadow-md animate-pulse' 
                    : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-300'
                }`}
                title="Toggle freehand draw mode so you can draw circle/note directly over cell elements"
              >
                {drawModeEnabled ? "✍ Drawing Mode: ON" : "✍ Enable Draw Pen"}
              </button>
            </div>

            {drawModeEnabled && (
              <>
                {/* Pen color picker circles */}
                <div className="flex items-center gap-2 border-r border-slate-200 pr-4">
                  <span className="text-[10px] uppercase font-bold text-slate-400">Pen Color:</span>
                  <div className="flex items-center gap-1 bg-white p-1 rounded-md border">
                    {[
                      { hex: '#e11d48', name: 'Rose Red' },
                      { hex: '#2563eb', name: 'Royal Blue' },
                      { hex: '#16a34a', name: 'Emerald Green' },
                      { hex: '#d97706', name: 'Amber Yellow' },
                      { hex: '#0f172a', name: 'Slate Black' }
                    ].map(pc => (
                      <button
                        key={pc.hex}
                        onClick={() => onDrawColorChange(pc.hex)}
                        className={`w-4 h-4 rounded-full transition-transform hover:scale-110 ${
                          drawColor === pc.hex ? 'ring-2 ring-slate-400 scale-105' : ''
                        }`}
                        style={{ backgroundColor: pc.hex }}
                        title={pc.name}
                      />
                    ))}
                  </div>
                </div>

                {/* Brush width slider */}
                <div className="flex items-center gap-2 border-r border-slate-200 pr-4">
                  <span className="text-[10px] uppercase font-bold text-slate-400">Brush Size: {drawThickness}px</span>
                  <input
                    type="range"
                    min="1"
                    max="8"
                    value={drawThickness}
                    onChange={(e) => onDrawThicknessChange(parseInt(e.target.value, 10))}
                    className="w-20 accent-rose-600 cursor-pointer"
                  />
                </div>
              </>
            )}

            {/* Clear sketches */}
            <button
              onClick={onClearDrawings}
              className="px-2.5 py-1 text-xs text-red-600 bg-white hover:bg-red-50 border border-red-200 rounded cursor-pointer font-medium ml-auto transition-colors"
              title="Delete all drawings sketch lines on this active sheet"
            >
              Clear Sketch Canvas
            </button>
          </div>
        )}

        {/* TAB 5: FORMULAS, DATA & SOURCE BACKUP */}
        {activeRibbonTab === 'formulas' && (
          <div className="flex flex-col gap-2 w-full">
            <div className="flex items-center flex-wrap gap-4 w-full">
              {/* Quick formulas inserts */}
              <div className="flex items-center gap-2 border-r border-slate-200 pr-4">
                <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                  <Sparkles className="w-3.5 h-3.5 text-slate-500" />
                  Math Formulas
                </span>
                <div className="flex flex-wrap gap-1">
                  {[
                    { name: 'SUM', template: '=SUM(A1:A5)', title: 'Add selection range values' },
                    { name: 'AVERAGE', template: '=AVERAGE(A1:A5)', title: 'Average range values' },
                    { name: 'MIN', template: '=MIN(A1:A5)', title: 'Get minimum value' },
                    { name: 'MAX', template: '=MAX(A1:A5)', title: 'Get maximum value' },
                    { name: 'COUNT', template: '=COUNT(A1:A5)', title: 'Count numerical records' }
                  ].map(fm => (
                    <button
                      key={fm.name}
                      onClick={() => onInsertFormulaValue(fm.template)}
                      className="px-2 py-0.5 bg-white hover:bg-slate-100 border border-slate-200 hover:border-slate-300 text-slate-700 text-xs rounded transition-all cursor-pointer font-medium"
                      title={fm.title}
                    >
                      {fm.name}
                    </button>
                  ))}
                </div>
              </div>

              {/* Quick date calculations */}
              <div className="flex items-center gap-2 border-r border-slate-200 pr-4">
                <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                  <Calendar className="w-3.5 h-3.5 text-slate-500" />
                  Calendar Dates
                </span>
                <div className="flex flex-wrap gap-1">
                  {[
                    { name: 'TODAY()', template: '=TODAY()', title: 'Insert current Date string "YYYY-MM-DD"' },
                    { name: 'NOW()', template: '=NOW()', title: 'Insert localized full Date & Time string' },
                    { name: 'DAYS', template: '=DAYS("2026-12-31", TODAY())', title: 'Days count between two date values' },
                    { name: 'DATEDIF', template: '=DATEDIF("1995-10-25", TODAY(), "Y")', title: 'Age/Difference in years ("Y"), months ("M"), or days ("D")' }
                  ].map(fm => (
                    <button
                      key={fm.name}
                      onClick={() => onInsertFormulaValue(fm.template)}
                      className={`px-2 py-0.5 bg-white hover:bg-slate-100 border border-slate-200 text-slate-700 text-xs rounded transition-all cursor-pointer font-medium`}
                      title={fm.title}
                    >
                      {fm.name}
                    </button>
                  ))}
                </div>
              </div>

              {/* Advanced Data Tools Group */}
              <div className="flex items-center gap-2 border-r border-slate-200 pr-4">
                <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                  <Database className="w-3.5 h-3.5 text-slate-500" />
                  Data Tools
                </span>
                <div className="flex items-center gap-1.5">
                  <button
                    onClick={onRemoveDuplicates}
                    className="px-2.5 py-1 text-xs font-semibold bg-white hover:bg-red-50 text-red-600 hover:text-red-700 border border-red-200 rounded transition-all cursor-pointer flex items-center gap-1"
                    title="Remove Duplicate rows inside your current selection range"
                  >
                    <Trash2 className="w-3 h-3" />
                    <span>Remove Duplicates</span>
                  </button>

                  <div className="flex items-center bg-white border border-slate-200 p-0.5 rounded-md gap-1">
                    <select
                      value={splitDelimiter}
                      onChange={(e) => setSplitDelimiter(e.target.value)}
                      className="bg-transparent border-0 text-[11px] font-bold focus:outline-hidden cursor-pointer text-slate-600"
                      title="Select splitting delimiter"
                    >
                      <option value=",">Comma (,)</option>
                      <option value="\t">Tab (\t)</option>
                      <option value=";">Semicolon (;)</option>
                      <option value=" ">Space ( )</option>
                    </select>
                    <button
                      onClick={() => onTextToColumns(splitDelimiter)}
                      className="px-2 py-0.5 text-[11px] font-bold bg-indigo-50 hover:bg-indigo-100 text-indigo-700 rounded cursor-pointer transition-all flex items-center gap-1"
                      title="Split cell content in active selection column by selected delimiter into separate columns"
                    >
                      <Columns4 className="w-3 h-3" />
                      <span>Text to Columns</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* Source Backup & Copy All Group */}
              <div className="flex items-center gap-2 border-r border-slate-200 pr-4">
                <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                  <Share2 className="w-3.5 h-3.5 text-indigo-500" />
                  Copy Source & Backup
                </span>
                <div className="flex items-center gap-1">
                  <button
                    onClick={onCopySheetCSVText}
                    className="px-2 py-1 text-xs bg-white hover:bg-slate-100 border border-slate-200 text-slate-700 rounded transition-all cursor-pointer font-medium flex items-center gap-1"
                    title="Copy current active sheet cell data as CSV table format to your clipboard instantly"
                  >
                    <CopyCheck className="w-3.5 h-3.5 text-green-600" />
                    <span>Copy CSV Table</span>
                  </button>
                  <button
                    onClick={onCopyWorkbookBackup}
                    className="px-2 py-1 text-xs bg-white hover:bg-slate-100 border border-slate-200 text-slate-700 rounded transition-all cursor-pointer font-semibold flex items-center gap-1"
                    title="Copy full Workbook structure (All sheets, values, formulas, styles, widgets, macros) to clipboard as a single backup JSON string"
                  >
                    <FileCode className="w-3.5 h-3.5 text-indigo-500" />
                    <span>Copy Backup JSON</span>
                  </button>
                  <button
                    onClick={() => setShowRestorePanel(!showRestorePanel)}
                    className={`px-2 py-1 text-xs border rounded transition-all cursor-pointer font-bold flex items-center gap-1 ${
                      showRestorePanel 
                        ? 'bg-indigo-100 border-indigo-300 text-indigo-800' 
                        : 'bg-indigo-600 hover:bg-indigo-700 border-indigo-500 text-white'
                    }`}
                    title="Open panel to paste and restore full workbook state"
                  >
                    <span>Restore Backup</span>
                  </button>
                </div>
              </div>

              {/* Show raw formulas toggle */}
              <div className="flex items-center gap-2 pl-2 ml-auto">
                <button
                  onClick={() => onShowFormulasActiveChange(!showFormulasActive)}
                  className={`px-3 py-1 text-xs font-bold rounded-lg cursor-pointer transition-colors ${
                    showFormulasActive 
                      ? 'bg-amber-600 hover:bg-amber-700 text-white shadow-md animate-pulse' 
                      : 'bg-white hover:bg-slate-100 text-slate-700 border border-slate-300'
                  }`}
                  title="Toggle displaying raw formulas instead of computed values"
                >
                  {showFormulasActive ? "👁️ Showing Formulas" : "👁️ Show Formulas"}
                </button>
              </div>
            </div>

            {/* EXPANDABLE WORKBOOK BACKUP RESTORE PANEL */}
            {showRestorePanel && (
              <div className="flex flex-col gap-1.5 p-3.5 bg-indigo-50/50 border border-indigo-100 rounded-lg w-full animate-fadeIn mt-1">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-bold text-indigo-900 flex items-center gap-1.5">
                    <Database className="w-4 h-4 text-indigo-600" />
                    Paste Full Workbook Source JSON Backup
                  </span>
                  <span className="text-[10px] text-indigo-500 italic">This will completely overwrite all sheets and restore drawing sketches, widgets, styles and values.</span>
                </div>
                <div className="flex gap-2">
                  <textarea
                    value={backupPasteText}
                    onChange={(e) => setBackupPasteText(e.target.value)}
                    placeholder="Paste your copied Workbook Backup JSON string here..."
                    className="w-full h-12 text-xs font-mono p-1.5 bg-white border border-indigo-200 rounded focus:outline-none focus:ring-1 focus:ring-indigo-400 resize-none"
                  />
                  <button
                    onClick={() => {
                      if (!backupPasteText.trim()) return;
                      onPasteWorkbookBackup(backupPasteText.trim());
                      setBackupPasteText("");
                      setShowRestorePanel(false);
                    }}
                    className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs rounded-md shadow-xs cursor-pointer flex-shrink-0 self-center"
                  >
                    Load Workbook Backup
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* TAB 6: AUTOMATIONS */}
        {activeRibbonTab === 'automate' && (
          <div className="flex items-center justify-between gap-4 w-full">
            <div className="flex flex-col gap-1 w-full max-w-lg">
              <div className="flex items-center justify-between">
                <span className="text-[10px] uppercase font-bold text-slate-400 flex items-center gap-1">
                  <Play className="w-3.5 h-3.5 text-emerald-500" />
                  Javascript Macros Console Engine
                </span>
                <div className="flex gap-2">
                  <button
                    onClick={() => setMacroScript(
                      `// Macro: Replicate values from A1 to range B1:B10\nfor (let r = 1; r <= 10; r++) {\n  let val = getCell("A1");\n  setCell("B" + r, val);\n}`
                    )}
                    className="text-[10px] text-slate-500 hover:text-slate-800 underline cursor-pointer"
                  >
                    Load Copy Template
                  </button>
                  <button
                    onClick={() => setMacroScript(
                      `// Macro: Highlight all rows with value > 100 in Column C\n// Hint: Set cells styles programmatically (requires custom style code in sheet or cell values)\nfor (let r = 1; r <= 30; r++) {\n  let val = parseFloat(getCell("C" + r) || 0);\n  if (val > 100) {\n    setCell("C" + r, "HIGH: " + val);\n  }\n}`
                    )}
                    className="text-[10px] text-slate-500 hover:text-slate-800 underline cursor-pointer"
                  >
                    Load Highlighting Template
                  </button>
                </div>
              </div>
              <textarea
                value={macroScript}
                onChange={(e) => setMacroScript(e.target.value)}
                rows={2}
                className="w-full text-[11px] font-mono bg-slate-900 text-green-400 p-2 rounded border border-slate-800 focus:outline-hidden resize-none"
                placeholder="Write custom javascript automation macro code..."
              />
            </div>

            <button
              onClick={() => onRunMacroScript(macroScript)}
              className="px-4 py-2 font-bold text-xs text-white bg-slate-900 hover:bg-slate-950 rounded-lg shadow-sm flex items-center gap-1.5 cursor-pointer flex-shrink-0"
              title="Execute Javascript Macro Code over active spreadsheet cells"
            >
              <Play className="w-3.5 h-3.5 text-emerald-400 fill-emerald-400" />
              <span>Run Script Macro</span>
            </button>
          </div>
        )}

      </div>

      {/* EXPANDABLE FIND AND REPLACE BLOCK */}
      {showFindReplace && (
        <div className="flex flex-wrap items-center gap-3 px-4 py-1.5 border-t border-slate-100 bg-slate-100/50">
          <div className="flex items-center gap-1.5">
            <span className="text-xs font-semibold text-slate-600">Find:</span>
            <input 
              type="text" 
              placeholder="Text or value to find"
              value={findText}
              onChange={(e) => setFindText(e.target.value)}
              className="px-2 py-0.5 text-xs bg-white border border-slate-300 rounded focus:outline-hidden min-w-[120px]"
            />
          </div>
          <div className="flex items-center gap-1.5">
            <span className="text-xs font-semibold text-slate-600">Replace:</span>
            <input 
              type="text" 
              placeholder="Replacement text"
              value={replaceText}
              onChange={(e) => setReplaceText(e.target.value)}
              className="px-2 py-0.5 text-xs bg-white border border-slate-300 rounded focus:outline-hidden min-w-[120px]"
            />
          </div>
          <button
            onClick={() => onFindReplace(findText, replaceText)}
            disabled={!findText}
            className={`px-2.5 py-0.5 text-xs font-bold text-white ${getThemeColorClass()} rounded disabled:opacity-40 cursor-pointer`}
          >
            Replace All
          </button>
          <span className="text-[10px] text-slate-400 italic">This will replace values across the entire active sheet.</span>
        </div>
      )}
    </div>
  );
};
