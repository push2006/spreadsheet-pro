import { useState, useMemo, useEffect } from 'react';
import { Toolbar } from './components/Toolbar';
import { FormulaBar } from './components/FormulaBar';
import { Grid } from './components/Grid';
import { Tabs } from './components/Tabs';
import { ImportModal } from './components/ImportModal';
import { Sheet, SheetData, SheetStyles, CellStyle, SelectionRange, WorkbookState, FloatingWidget } from './types';
import { getCellEvaluatedValue, parseCellId, indexToColName, colNameToIndex } from './utils/formulaEvaluator';
import { FileSpreadsheet, HelpCircle, X, AlertCircle, Info } from 'lucide-react';

// Generates an empty default grid
function createInitialWorkbook(): WorkbookState {
  const initialData: SheetData = {};
  const initialStyles: SheetStyles = {};

  const firstSheet: Sheet = {
    id: "sheet_1",
    name: "Sheet 1",
    data: initialData,
    styles: initialStyles,
    columnWidths: {},
    rowHeights: {},
    mergedCells: [],
    hiddenColumns: {},
    hiddenRows: {}
  };

  return {
    sheets: { [firstSheet.id]: firstSheet },
    sheetIds: [firstSheet.id],
    activeSheetId: firstSheet.id
  };
}

// Robust validator to prevent corrupt localStorage states from crashing the app
function validateAndCleanWorkbook(parsed: any): WorkbookState | null {
  try {
    if (!parsed || typeof parsed !== 'object') return null;
    if (!parsed.sheets || typeof parsed.sheets !== 'object') return null;
    if (!Array.isArray(parsed.sheetIds) || parsed.sheetIds.length === 0) return null;
    
    // Ensure activeSheetId is set and valid
    let activeId = parsed.activeSheetId;
    if (typeof activeId !== 'string' || !parsed.sheets[activeId]) {
      activeId = parsed.sheetIds[0];
    }
    if (!parsed.sheets[activeId]) return null;

    // Validate and clean each sheet
    const cleanSheets: { [id: string]: Sheet } = {};
    const cleanSheetIds: string[] = [];

    parsed.sheetIds.forEach((id: string) => {
      const sheet = parsed.sheets[id];
      if (sheet && typeof sheet === 'object') {
        cleanSheets[id] = {
          id: sheet.id || id,
          name: sheet.name || `Sheet ${id}`,
          data: (sheet.data && typeof sheet.data === 'object') ? sheet.data : {},
          styles: (sheet.styles && typeof sheet.styles === 'object') ? sheet.styles : {},
          columnWidths: (sheet.columnWidths && typeof sheet.columnWidths === 'object') ? sheet.columnWidths : {},
          rowHeights: (sheet.rowHeights && typeof sheet.rowHeights === 'object') ? sheet.rowHeights : {},
          mergedCells: Array.isArray(sheet.mergedCells) ? sheet.mergedCells : [],
          hiddenColumns: (sheet.hiddenColumns && typeof sheet.hiddenColumns === 'object') ? sheet.hiddenColumns : {},
          hiddenRows: (sheet.hiddenRows && typeof sheet.hiddenRows === 'object') ? sheet.hiddenRows : {}
        };
        cleanSheetIds.push(id);
      }
    });

    if (cleanSheetIds.length === 0) return null;

    return {
      sheets: cleanSheets,
      sheetIds: cleanSheetIds,
      activeSheetId: cleanSheets[activeId] ? activeId : cleanSheetIds[0]
    };
  } catch (e) {
    console.error("Failed to validate saved workbook:", e);
    return null;
  }
}

export default function App() {
  const [workbook, setWorkbook] = useState<WorkbookState>(() => {
    try {
      const saved = localStorage.getItem("spreadsheetWorkbook_v3");
      if (saved) {
        const parsed = JSON.parse(saved);
        const validated = validateAndCleanWorkbook(parsed);
        if (validated) {
          return validated;
        }
      }
    } catch (e) {
      console.warn("Could not read saved spreadsheet, building new default", e);
    }
    return createInitialWorkbook();
  });

  const [selectedCellId, setSelectedCellId] = useState<string | null>("A1");
  const [activeRange, setActiveRange] = useState<SelectionRange | null>(null);
  const [isImportOpen, setIsImportOpen] = useState(false);
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [isHelpOpen, setIsHelpOpen] = useState(false);
  const [showFormulasActive, setShowFormulasActive] = useState<boolean>(false);

  // Advanced features and workbook customization states
  const [themeColor, setThemeColor] = useState<'green' | 'blue' | 'purple' | 'rose' | 'amber' | 'slate'>(() => {
    return (localStorage.getItem("spreadsheet_themeColor") as any) || "green";
  });
  const [hideGridlines, setHideGridlines] = useState<boolean>(() => {
    return localStorage.getItem("spreadsheet_hideGridlines") === "true";
  });
  const [drawModeEnabled, setDrawModeEnabled] = useState<boolean>(false);
  const [drawColor, setDrawColor] = useState<string>("#e11d48"); // default draw color (rose red)
  const [drawThickness, setDrawThickness] = useState<number>(3);
  const [drawLines, setDrawLines] = useState<{ [sheetId: string]: Array<{ points: number[]; color: string; width: number }> }>(() => {
    try {
      const saved = localStorage.getItem("spreadsheet_drawLines");
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });
  const [floatingWidgets, setFloatingWidgets] = useState<FloatingWidget[]>(() => {
    try {
      const saved = localStorage.getItem("spreadsheet_floatingWidgets");
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  // Sync state modifications to persistence store
  useEffect(() => {
    localStorage.setItem("spreadsheet_themeColor", themeColor);
  }, [themeColor]);

  useEffect(() => {
    localStorage.setItem("spreadsheet_hideGridlines", String(hideGridlines));
  }, [hideGridlines]);

  // Auto-dismiss toast after 3 seconds
  useEffect(() => {
    if (!toastMessage) return;
    const timer = setTimeout(() => setToastMessage(null), 3000);
    return () => clearTimeout(timer);
  }, [toastMessage]);

  useEffect(() => {
    localStorage.setItem("spreadsheet_drawLines", JSON.stringify(drawLines));
  }, [drawLines]);

  useEffect(() => {
    localStorage.setItem("spreadsheet_floatingWidgets", JSON.stringify(floatingWidgets));
  }, [floatingWidgets]);

  // Dynamic Theme Colors Mapper
  const themeStyles = useMemo(() => {
    switch (themeColor) {
      case 'blue':
        return {
          bgHeader: 'bg-blue-700',
          bgButton: 'bg-blue-600 hover:bg-blue-700',
          bgHeaderAccent: 'bg-blue-800 hover:bg-blue-900 border-blue-600/50',
          textAccent: 'text-blue-200',
          textSub: 'text-blue-100',
          accentColor: '#2563eb',
          ringColor: 'focus:ring-blue-500',
          borderColor: 'border-blue-500',
          textBadge: 'text-blue-300',
          textTheme: 'text-blue-600'
        };
      case 'purple':
        return {
          bgHeader: 'bg-purple-700',
          bgButton: 'bg-purple-600 hover:bg-purple-700',
          bgHeaderAccent: 'bg-purple-800 hover:bg-purple-900 border-purple-600/50',
          textAccent: 'text-purple-200',
          textSub: 'text-purple-100',
          accentColor: '#9333ea',
          ringColor: 'focus:ring-purple-500',
          borderColor: 'border-purple-500',
          textBadge: 'text-purple-300',
          textTheme: 'text-purple-600'
        };
      case 'rose':
        return {
          bgHeader: 'bg-rose-700',
          bgButton: 'bg-rose-600 hover:bg-rose-700',
          bgHeaderAccent: 'bg-rose-800 hover:bg-rose-900 border-rose-600/50',
          textAccent: 'text-rose-200',
          textSub: 'text-rose-100',
          accentColor: '#e11d48',
          ringColor: 'focus:ring-rose-500',
          borderColor: 'border-rose-500',
          textBadge: 'text-rose-300',
          textTheme: 'text-rose-600'
        };
      case 'amber':
        return {
          bgHeader: 'bg-amber-700',
          bgButton: 'bg-amber-600 hover:bg-amber-700',
          bgHeaderAccent: 'bg-amber-800 hover:bg-amber-900 border-amber-600/50',
          textAccent: 'text-amber-200',
          textSub: 'text-amber-100',
          accentColor: '#d97706',
          ringColor: 'focus:ring-amber-500',
          borderColor: 'border-amber-500',
          textBadge: 'text-amber-300',
          textTheme: 'text-amber-600'
        };
      case 'slate':
        return {
          bgHeader: 'bg-slate-800',
          bgButton: 'bg-slate-700 hover:bg-slate-800',
          bgHeaderAccent: 'bg-slate-900 hover:bg-slate-950 border-slate-700/50',
          textAccent: 'text-slate-200',
          textSub: 'text-slate-300',
          accentColor: '#475569',
          ringColor: 'focus:ring-slate-500',
          borderColor: 'border-slate-500',
          textBadge: 'text-slate-400',
          textTheme: 'text-slate-700'
        };
      case 'green':
      default:
        return {
          bgHeader: 'bg-green-700',
          bgButton: 'bg-green-600 hover:bg-green-700',
          bgHeaderAccent: 'bg-green-800 hover:bg-green-900 border-green-600/50',
          textAccent: 'text-green-200',
          textSub: 'text-green-100',
          accentColor: '#16a34a',
          ringColor: 'focus:ring-green-500',
          borderColor: 'border-green-500',
          textBadge: 'text-green-300',
          textTheme: 'text-green-600'
        };
    }
  }, [themeColor]);

  // Cell Name Box Input value
  const [cellNameInput, setCellNameInput] = useState("A1");

  // Clipboard context state
  const [clipboard, setClipboard] = useState<{
    type: 'copy' | 'cut';
    cells: { colOffset: number; rowOffset: number; value: string; style: CellStyle | null }[];
    sourceSheetId: string;
    sourceRange: SelectionRange;
  } | null>(null);

  // Format Painter state: if active, contains the copied CellStyle
  const [formatPainterStyle, setFormatPainterStyle] = useState<CellStyle | null>(null);

  // Undo / Redo histories snapshot vectors
  const [undoStack, setUndoStack] = useState<WorkbookState[]>([]);
  const [redoStack, setRedoStack] = useState<WorkbookState[]>([]);

  // Sync CellNameInput with SelectedCellId
  useEffect(() => {
    if (selectedCellId) {
      setCellNameInput(selectedCellId);
    }
  }, [selectedCellId]);

  // Calculate current sheet safely
  const activeSheet = useMemo(() => {
    if (!workbook || !workbook.sheets || !workbook.activeSheetId) return null;
    return workbook.sheets[workbook.activeSheetId] || Object.values(workbook.sheets)[0] || null;
  }, [workbook]);

  // Persist workbook to LocalStorage whenever changes happen to workbook state
  useEffect(() => {
    if (workbook) {
      localStorage.setItem("spreadsheetWorkbook_v3", JSON.stringify(workbook));
    }
  }, [workbook]);

  // Automatically recalculate evaluated values whenever activeSheet.data undergoes changes
  const evaluatedValues = useMemo(() => {
    if (!activeSheet || !activeSheet.data) return {};
    const cache: { [cellId: string]: string } = {};
    const keys = Object.keys(activeSheet.data);
    
    // Evaluate cell by cell. Interdependent logic resolved recursively
    keys.forEach(cellId => {
      getCellEvaluatedValue(cellId, activeSheet.data, cache);
    });
    
    return cache;
  }, [activeSheet]);

  // Read current selected cell raw entered string (e.g. "=SUM(B5:B10)" or "Rent")
  const selectedCellFormula = useMemo(() => {
    if (!selectedCellId || !activeSheet || !activeSheet.data) return "";
    return activeSheet.data[selectedCellId] || "";
  }, [selectedCellId, activeSheet]);

  // Push current workbook to history stack before committing edits
  const pushToHistory = (currentState: WorkbookState = workbook) => {
    const cloned = JSON.parse(JSON.stringify(currentState));
    setUndoStack(prev => [...prev.slice(-30), cloned]); // Limit stack to 30 actions
    setRedoStack([]); // reset redo stack on new action
  };

  // Undo execution handler
  const handleUndo = () => {
    if (undoStack.length === 0) return;
    const previous = undoStack[undoStack.length - 1];
    setUndoStack(prev => prev.slice(0, -1));

    const clonedCurrent = JSON.parse(JSON.stringify(workbook));
    setRedoStack(prev => [...prev, clonedCurrent]);

    setWorkbook(previous);
  };

  // Redo execution handler
  const handleRedo = () => {
    if (redoStack.length === 0) return;
    const next = redoStack[redoStack.length - 1];
    setRedoStack(prev => prev.slice(0, -1));

    const clonedCurrent = JSON.parse(JSON.stringify(workbook));
    setUndoStack(prev => [...prev, clonedCurrent]);

    setWorkbook(next);
  };

  const applyFormatPainting = (targetRange: SelectionRange | null, targetCell: string | null) => {
    if (!formatPainterStyle) return;
    
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      if (!sheet.styles) sheet.styles = {};
      
      let cellsToPaint: string[] = [];
      if (targetRange) {
        const startParsed = parseCellId(targetRange.start);
        const endParsed = parseCellId(targetRange.end);
        if (startParsed && endParsed) {
          const minCol = Math.min(startParsed.colIndex, endParsed.colIndex);
          const maxCol = Math.max(startParsed.colIndex, endParsed.colIndex);
          const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
          const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);
          for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
              cellsToPaint.push(indexToColName(c) + r);
            }
          }
        }
      } else if (targetCell) {
        cellsToPaint.push(targetCell);
      }
      
      cellsToPaint.forEach(cellId => {
        sheet.styles[cellId] = { ...formatPainterStyle };
      });
      
      return cloned;
    });
    
    setFormatPainterStyle(null);
    setToastMessage("Format painted successfully!");
  };

  const handleToggleFormatPainter = () => {
    if (formatPainterStyle) {
      setFormatPainterStyle(null);
      setToastMessage("Format Painter deactivated");
    } else {
      if (!selectedCellId) {
        setToastMessage("Please select a cell first to copy its format");
        return;
      }
      setFormatPainterStyle(currentCellStyle);
      setToastMessage("Format Painter activated! Click or drag select cells to paint format");
    }
  };

  // Switch selected cell
  const handleSelectCell = (cellId: string) => {
    if (formatPainterStyle) {
      applyFormatPainting(null, cellId);
    }

    const formulaInput = document.getElementById("formula-bar-input") as HTMLInputElement | null;
    const isFormulaBarFocused = formulaInput && document.activeElement === formulaInput;
    const currentVal = workbook.sheets[workbook.activeSheetId].data[selectedCellId || ""] || "";
    
    if (isFormulaBarFocused && (currentVal.startsWith("=") || currentVal.startsWith("+"))) {
      // Append coordinates at current cursor pos of formula bar
      const start = formulaInput.selectionStart ?? currentVal.length;
      const end = formulaInput.selectionEnd ?? currentVal.length;
      const newVal = currentVal.slice(0, start) + cellId + currentVal.slice(end);
      handleCellChanged(selectedCellId!, newVal);
      
      // Keep focus on formula bar input
      setTimeout(() => {
        formulaInput.focus();
        const newPos = start + cellId.length;
        formulaInput.setSelectionRange(newPos, newPos);
      }, 0);
      return;
    }

    setSelectedCellId(cellId);
    setActiveRange(null); // clear range selection
  };

  // Coordinate jumped selector
  const handleCellNameInputSubmit = () => {
    const clean = cellNameInput.trim().toUpperCase();
    const parsed = parseCellId(clean);
    if (parsed) {
      setSelectedCellId(clean);
      setActiveRange(null);
    } else {
      setToastMessage("Invalid cell coordinate (e.g. B10, A1)");
      if (selectedCellId) {
        setCellNameInput(selectedCellId);
      }
    }
  };

  const handleImportData = (importedData: { [cellId: string]: string }, startCell: string) => {
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      if (!sheet.data) sheet.data = {};
      
      Object.entries(importedData).forEach(([cellId, value]) => {
        if (value.trim() === "") {
          delete sheet.data[cellId];
        } else {
          sheet.data[cellId] = value;
        }
      });
      return cloned;
    });
    setSelectedCellId(startCell);
  };

  // Drag selection handlers
  const handleStartRangeSelection = (cellId: string) => {
    setActiveRange({ start: cellId, end: cellId });
  };

  const handleExpandRangeSelection = (cellId: string) => {
    setActiveRange(prev => {
      if (!prev) return { start: cellId, end: cellId };
      return { ...prev, end: cellId };
    });
  };

  const handleEndRangeSelection = () => {
    if (formatPainterStyle && activeRange) {
      applyFormatPainting(activeRange, null);
    }
  };

  // Set individual or mass cells value updates
  const handleCellChanged = (cellId: string, newValue: string) => {
    pushToHistory();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      if (!sheet.data) sheet.data = {};
      
      if (newValue.trim() === "") {
        delete sheet.data[cellId];
      } else {
        sheet.data[cellId] = newValue;
      }
      return cloned;
    });
  };

  // Update styles for highlighted cell or entire highlighted range of cells!
  const handleUpdateStyle = (styleUpdates: Partial<CellStyle>) => {
    pushToHistory();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      if (!sheet.styles) sheet.styles = {};
      
      const targetCells = getActiveSelectedCellIds();
      
      targetCells.forEach(cellId => {
        const existing = sheet.styles[cellId] || {};
        sheet.styles[cellId] = {
          ...existing,
          ...styleUpdates
        };
      });

      return cloned;
    });
  };

  // Fetch every cellId currently selected (either single or multi-range)
  const getActiveSelectedCellIds = (): string[] => {
    if (!selectedCellId) return [];
    if (!activeRange) return [selectedCellId];

    const startParsed = parseCellId(activeRange.start);
    const endParsed = parseCellId(activeRange.end);
    if (!startParsed || !endParsed) return [selectedCellId];

    const minCol = Math.min(startParsed.colIndex, endParsed.colIndex);
    const maxCol = Math.max(startParsed.colIndex, endParsed.colIndex);
    const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
    const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);

    const rangeIds: string[] = [];
    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        rangeIds.push(indexToColName(c) + r);
      }
    }
    return rangeIds;
  };

  // Helper range calculations
  const getCellsInRange = (start: string, end: string) => {
    const startParsed = parseCellId(start);
    const endParsed = parseCellId(end);
    if (!startParsed || !endParsed) return [];
    
    const minCol = Math.min(startParsed.colIndex, endParsed.colIndex);
    const maxCol = Math.max(startParsed.colIndex, endParsed.colIndex);
    const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
    const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);
    
    const cells: { cellId: string; colOffset: number; rowOffset: number }[] = [];
    for (let c = minCol; c <= maxCol; c++) {
      for (let r = minRow; r <= maxRow; r++) {
        const cellId = indexToColName(c) + r;
        cells.push({
          cellId,
          colOffset: c - minCol,
          rowOffset: r - minRow
        });
      }
    }
    return cells;
  };

  // Clipboard Copier
  const handleCopy = () => {
    if (!selectedCellId || !activeSheet) return;
    const activeSheetData = activeSheet.data;
    const activeSheetStyles = activeSheet.styles;
    
    let clipboardCells: { colOffset: number; rowOffset: number; value: string; style: CellStyle | null }[] = [];
    
    if (activeRange) {
      const cells = getCellsInRange(activeRange.start, activeRange.end);
      cells.forEach(({ cellId, colOffset, rowOffset }) => {
        clipboardCells.push({
          colOffset,
          rowOffset,
          value: activeSheetData[cellId] || "",
          style: activeSheetStyles[cellId] || null
        });
      });
    } else {
      clipboardCells.push({
        colOffset: 0,
        rowOffset: 0,
        value: activeSheetData[selectedCellId] || "",
        style: activeSheetStyles[selectedCellId] || null
      });
    }
    
    setClipboard({
      type: 'copy',
      cells: clipboardCells,
      sourceSheetId: workbook.activeSheetId,
      sourceRange: activeRange ? { ...activeRange } : { start: selectedCellId, end: selectedCellId }
    });
    
    setToastMessage("Selected cells copied to clipboard");
  };

  // Clipboard Cutter
  const handleCut = () => {
    if (!selectedCellId || !activeSheet) return;
    const activeSheetData = activeSheet.data;
    const activeSheetStyles = activeSheet.styles;
    
    let clipboardCells: { colOffset: number; rowOffset: number; value: string; style: CellStyle | null }[] = [];
    
    let rangeToCut = activeRange ? { ...activeRange } : { start: selectedCellId, end: selectedCellId };
    const cells = getCellsInRange(rangeToCut.start, rangeToCut.end);
    cells.forEach(({ cellId, colOffset, rowOffset }) => {
      clipboardCells.push({
        colOffset,
        rowOffset,
        value: activeSheetData[cellId] || "",
        style: activeSheetStyles[cellId] || null
      });
    });
    
    setClipboard({
      type: 'cut',
      cells: clipboardCells,
      sourceSheetId: workbook.activeSheetId,
      sourceRange: rangeToCut
    });
    
    setToastMessage("Selected cells cut to clipboard");
  };

  // Clipboard Paster with Paste Special options
  const handlePasteSpecial = (mode: 'all' | 'values' | 'formats' | 'formulas' | 'transpose') => {
    if (!selectedCellId || !clipboard) return;
    
    pushToHistory();
    setWorkbook(prev => {
      const activeId = prev.activeSheetId;
      const sheet = prev.sheets[activeId];
      const newData = { ...sheet.data };
      const newStyles = { ...sheet.styles };
      
      const targetParsed = parseCellId(selectedCellId);
      if (!targetParsed) return prev;
      
      const startColIdx = targetParsed.colIndex;
      const startRowIdx = targetParsed.rowIndex;
      
      // Paste data
      clipboard.cells.forEach(item => {
        let targetColName = "";
        let targetRowName = 0;

        if (mode === 'transpose') {
          targetColName = indexToColName(startColIdx + item.rowOffset);
          targetRowName = startRowIdx + item.colOffset;
        } else {
          targetColName = indexToColName(startColIdx + item.colOffset);
          targetRowName = startRowIdx + item.rowOffset;
        }

        const targetCellId = targetColName + targetRowName;
        
        if (mode === 'all' || mode === 'values' || mode === 'formulas' || mode === 'transpose') {
          if (item.value !== "") {
            newData[targetCellId] = item.value;
          } else {
            delete newData[targetCellId];
          }
        }
        
        if (mode === 'all' || mode === 'formats' || mode === 'transpose') {
          if (item.style) {
            newStyles[targetCellId] = item.style;
          } else {
            delete newStyles[targetCellId];
          }
        }
      });
      
      // Clear cut sources
      if (clipboard.type === 'cut') {
        const sourceCells = getCellsInRange(clipboard.sourceRange.start, clipboard.sourceRange.end);
        
        if (clipboard.sourceSheetId === activeId) {
          sourceCells.forEach(({ cellId }) => {
            if (mode === 'all' || mode === 'values' || mode === 'formulas' || mode === 'transpose') {
              delete newData[cellId];
            }
            if (mode === 'all' || mode === 'formats' || mode === 'transpose') {
              delete newStyles[cellId];
            }
          });
          
          // Re-apply pasted values on top of cuts
          clipboard.cells.forEach(item => {
            let targetColName = "";
            let targetRowName = 0;
            if (mode === 'transpose') {
              targetColName = indexToColName(startColIdx + item.rowOffset);
              targetRowName = startRowIdx + item.colOffset;
            } else {
              targetColName = indexToColName(startColIdx + item.colOffset);
              targetRowName = startRowIdx + item.rowOffset;
            }
            const targetCellId = targetColName + targetRowName;
            if (mode === 'all' || mode === 'values' || mode === 'formulas' || mode === 'transpose') {
              if (item.value !== "") {
                newData[targetCellId] = item.value;
              }
            }
            if (mode === 'all' || mode === 'formats' || mode === 'transpose') {
              if (item.style) {
                newStyles[targetCellId] = item.style;
              }
            }
          });
        } else {
          // Cut from another sheet
          const sourceSheet = prev.sheets[clipboard.sourceSheetId];
          const sourceData = { ...sourceSheet.data };
          const sourceStyles = { ...sourceSheet.styles };
          sourceCells.forEach(({ cellId }) => {
            if (mode === 'all' || mode === 'values' || mode === 'formulas' || mode === 'transpose') {
              delete sourceData[cellId];
            }
            if (mode === 'all' || mode === 'formats' || mode === 'transpose') {
              delete sourceStyles[cellId];
            }
          });
          
          prev.sheets[clipboard.sourceSheetId] = {
            ...sourceSheet,
            data: sourceData,
            styles: sourceStyles
          };
        }
      }
      
      return {
        ...prev,
        sheets: {
          ...prev.sheets,
          [activeId]: {
            ...sheet,
            data: newData,
            styles: newStyles
          }
        }
      };
    });
    
    if (clipboard.type === 'cut') {
      setClipboard(null);
    }
    const modeLabel = mode === 'all' ? 'cells' : mode === 'values' ? 'values only' : mode === 'formats' ? 'formats only' : mode === 'formulas' ? 'formulas only' : 'transposed cells';
    setToastMessage(`Selected ${modeLabel} pasted successfully`);
  };

  const handlePaste = () => {
    handlePasteSpecial('all');
  };

  // Row and Column Hide / Unhide mechanics
  const handleHideColumn = (col: string) => {
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet.hiddenColumns) sheet.hiddenColumns = {};
      sheet.hiddenColumns[col] = true;
      return cloned;
    });
    setToastMessage(`Column ${col} has been hidden`);
  };

  const handleUnhideAllColumns = () => {
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      sheet.hiddenColumns = {};
      return cloned;
    });
    setToastMessage("All columns have been unhidden");
  };

  const handleHideRow = (row: number) => {
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet.hiddenRows) sheet.hiddenRows = {};
      sheet.hiddenRows[row] = true;
      return cloned;
    });
    setToastMessage(`Row ${row} has been hidden`);
  };

  const handleUnhideAllRows = () => {
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      sheet.hiddenRows = {};
      return cloned;
    });
    setToastMessage("All rows have been unhidden");
  };

  // Sort targeted column rows alphabetically
  const handleSortActiveCol = (ascending: boolean) => {
    if (!selectedCellId) return;
    const parsed = parseCellId(selectedCellId);
    if (!parsed) return;

    pushToHistory();
    const colName = parsed.col;

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      if (!sheet.data) sheet.data = {};
      if (!sheet.styles) sheet.styles = {};
      
      const MAX_ROWS_TO_SORT = 50;
      const items: { row: number; val: string; formula: string; style: CellStyle }[] = [];
      for (let r = 1; r <= MAX_ROWS_TO_SORT; r++) { 
        const cellId = colName + r;
        items.push({
          row: r,
          val: evaluatedValues[cellId] || "",
          formula: sheet.data[cellId] || "",
          style: sheet.styles[cellId] || {}
        });
      }

      const filled = items.filter(i => i.formula.trim() !== "");
      const empty = items.filter(i => i.formula.trim() === "");

      filled.sort((a, b) => {
        const checkA = isNaN(Number(a.val)) ? a.val : Number(a.val);
        const checkB = isNaN(Number(b.val)) ? b.val : Number(b.val);

        if (typeof checkA === 'number' && typeof checkB === 'number') {
          return ascending ? checkA - checkB : checkB - checkA;
        }
        
        const strA = String(checkA).toLowerCase();
        const strB = String(checkB).toLowerCase();
        return ascending ? strA.localeCompare(strB) : strB.localeCompare(strA);
      });

      const combined = [...filled, ...empty];

      combined.forEach((item, index) => {
        const targetRow = index + 1; 
        const targetCellId = colName + targetRow;
        
        if (item.formula === "") {
          delete sheet.data[targetCellId];
          delete sheet.styles[targetCellId];
        } else {
          sheet.data[targetCellId] = item.formula;
          sheet.styles[targetCellId] = item.style;
        }
      });

      return cloned;
    });
  };

  // Mass Autofill replicate
  const handleAutoFillSelection = () => {
    if (!selectedCellId || !activeRange || !activeSheet) return;
    const baseVal = activeSheet.data[selectedCellId] || "";
    const baseStyle = activeSheet.styles[selectedCellId] || {};

    pushToHistory();
    const cells = getActiveSelectedCellIds();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      if (!sheet.data) sheet.data = {};
      if (!sheet.styles) sheet.styles = {};

      cells.forEach(cellId => {
        if (baseVal === "") {
          delete sheet.data[cellId];
        } else {
          sheet.data[cellId] = baseVal;
        }
        sheet.styles[cellId] = { ...baseStyle };
      });

      return cloned;
    });
  };

  // Find & Replace matches
  const handleFindReplace = (findText: string, replaceText: string) => {
    if (!findText) return;
    pushToHistory();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      if (!sheet.data) sheet.data = {};
      
      Object.keys(sheet.data).forEach(cellId => {
        const rawStr = sheet.data[cellId];
        if (rawStr && rawStr.includes(findText)) {
          sheet.data[cellId] = rawStr.replaceAll(findText, replaceText);
        }
      });

      return cloned;
    });
  };

  // CSV Generator downloader
  const handleExportCSV = () => {
    if (!activeSheet) return;
    
    // Dynamically compute used range from actual data
    const cellKeys = Object.keys(activeSheet.data);
    let maxCol = 0;
    let maxRow = 0;
    cellKeys.forEach(cellId => {
      const parsed = parseCellId(cellId);
      if (parsed) {
        if (parsed.colIndex > maxCol) maxCol = parsed.colIndex;
        if (parsed.rowIndex > maxRow) maxRow = parsed.rowIndex;
      }
    });
    // At least one page worth of rows/cols even if empty
    maxCol = Math.max(maxCol, 14);
    maxRow = Math.max(maxRow, 34);

    let csvContent = "";
    const cols = Array.from({ length: maxCol + 1 }, (_, i) => indexToColName(i)); 
    
    for (let r = 1; r <= maxRow + 1; r++) {
      const lineValues = cols.map(col => {
        const cellId = col + r;
        const val = evaluatedValues[cellId] !== undefined ? evaluatedValues[cellId] : (activeSheet.data[cellId] || "");
        const escaped = String(val).replace(/"/g, '""');
        return `"${escaped}"`;
      });
      csvContent += lineValues.join(",") + "\r\n";
    }

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", `${activeSheet.name.replace(/\s+/g, '_')}_export.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Add sheet to workbook state
  const handleAddSheet = () => {
    pushToHistory();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const count = cloned.sheetIds.length + 1;
      const newId = `sheet_${Date.now()}`;
      const newSheet: Sheet = {
        id: newId,
        name: `Sheet ${count}`,
        data: {},
        styles: {},
        columnWidths: {},
        rowHeights: {},
        mergedCells: [],
        hiddenColumns: {},
        hiddenRows: {}
      };

      cloned.sheets[newId] = newSheet;
      cloned.sheetIds.push(newId);
      cloned.activeSheetId = newId;

      return cloned;
    });
    
    setSelectedCellId("A1");
    setActiveRange(null);
  };

  // Duplicate currently active sheet (Excel/GSheets Copy Sheet)
  const handleDuplicateSheet = () => {
    if (!activeSheet) return;
    pushToHistory();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sourceSheet = cloned.sheets[cloned.activeSheetId];
      const newId = `sheet_${Date.now()}`;
      
      const duplicatedSheet: Sheet = {
        ...sourceSheet,
        id: newId,
        name: `${sourceSheet.name} (Copy)`,
      };

      cloned.sheets[newId] = duplicatedSheet;
      cloned.sheetIds.push(newId);
      cloned.activeSheetId = newId;

      return cloned;
    });

    setToastMessage(`Duplicated sheet "${activeSheet.name}" successfully!`);
    setSelectedCellId("A1");
    setActiveRange(null);
  };

  // Rename currently active sheet
  const handleRenameSheet = (newName: string) => {
    pushToHistory();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      if (cloned.sheets[cloned.activeSheetId]) {
        cloned.sheets[cloned.activeSheetId].name = newName;
      }
      return cloned;
    });
  };

  // Delete current active sheet
  const handleDeleteSheet = () => {
    if (workbook.sheetIds.length <= 1) {
      setToastMessage("Workbook must contain at least one worksheet sheet tab.");
      return;
    }

    pushToHistory();

    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const activeIdx = cloned.sheetIds.indexOf(cloned.activeSheetId);
      
      delete cloned.sheets[cloned.activeSheetId];
      cloned.sheetIds.splice(activeIdx, 1);
      
      const targetIdx = Math.max(0, activeIdx - 1);
      cloned.activeSheetId = cloned.sheetIds[targetIdx];

      return cloned;
    });

    setSelectedCellId("A1");
    setActiveRange(null);
  };

  // Copy full workbook JSON backup
  const handleCopyWorkbookBackup = () => {
    try {
      const backupObj = {
        version: "1.0",
        workbook,
        drawLines
      };
      const jsonStr = JSON.stringify(backupObj, null, 2);
      navigator.clipboard.writeText(jsonStr);
      setToastMessage("Full Workbook backup JSON successfully copied to clipboard!");
    } catch (err) {
      setToastMessage("Failed to copy backup JSON. Check browser permissions.");
    }
  };

  // Restore full workbook JSON backup
  const handlePasteWorkbookBackup = (jsonString: string) => {
    try {
      const parsed = JSON.parse(jsonString);
      if (!parsed || typeof parsed !== 'object') {
        setToastMessage("Invalid backup format!");
        return;
      }
      pushToHistory();
      if (parsed.workbook && parsed.workbook.sheets && parsed.workbook.activeSheetId) {
        setWorkbook(parsed.workbook);
        if (parsed.drawLines) {
          setDrawLines(parsed.drawLines);
        }
        setToastMessage("Full workbook and drawings restored successfully!");
      } else {
        setToastMessage("Malformed workbook state JSON.");
      }
    } catch (err) {
      setToastMessage("Failed to parse backup string. Make sure it is valid JSON.");
    }
  };

  // Copy active sheet CSV table text directly to clipboard
  const handleCopySheetCSVText = () => {
    if (!activeSheet) return;
    try {
      const cellKeys = Object.keys(activeSheet.data);
      let maxCol = 14;
      let maxRow = 34;
      cellKeys.forEach(cellId => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          if (parsed.colIndex > maxCol) maxCol = parsed.colIndex;
          if (parsed.rowIndex > maxRow) maxRow = parsed.rowIndex;
        }
      });

      let csvContent = "";
      const cols = Array.from({ length: maxCol + 1 }, (_, i) => indexToColName(i)); 
      for (let r = 1; r <= maxRow + 1; r++) {
        const lineValues = cols.map(col => {
          const cellId = col + r;
          const val = evaluatedValues[cellId] !== undefined ? evaluatedValues[cellId] : (activeSheet.data[cellId] || "");
          const escaped = String(val).replace(/"/g, '""');
          return `"${escaped}"`;
        });
        csvContent += lineValues.join(",") + "\n";
      }
      navigator.clipboard.writeText(csvContent);
      setToastMessage("Active sheet CSV table data copied to system clipboard!");
    } catch (err) {
      setToastMessage("Failed to copy CSV data to clipboard.");
    }
  };

  // Remove duplicates in selection range
  const handleRemoveDuplicates = () => {
    if (!activeSheet || !activeRange) {
      setToastMessage("Please select a range of multiple rows to remove duplicates.");
      return;
    }
    const startParsed = parseCellId(activeRange.start);
    const endParsed = parseCellId(activeRange.end);
    if (!startParsed || !endParsed) return;

    const minCol = Math.min(startParsed.colIndex, endParsed.colIndex);
    const maxCol = Math.max(startParsed.colIndex, endParsed.colIndex);
    const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
    const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);

    if (minRow === maxRow) {
      setToastMessage("Please select a range spanning multiple rows to check for duplicates.");
      return;
    }

    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;

      const rowSignatures: { [signature: string]: boolean } = {};
      const duplicateRows: number[] = [];

      for (let r = minRow; r <= maxRow; r++) {
        const rowValues: string[] = [];
        for (let c = minCol; c <= maxCol; c++) {
          const cellId = indexToColName(c) + r;
          rowValues.push(sheet.data[cellId] || "");
        }
        const signature = JSON.stringify(rowValues);
        if (rowSignatures[signature]) {
          duplicateRows.push(r);
        } else {
          rowSignatures[signature] = true;
        }
      }

      if (duplicateRows.length === 0) {
        return prev;
      }

      duplicateRows.forEach(r => {
        for (let c = minCol; c <= maxCol; c++) {
          const cellId = indexToColName(c) + r;
          delete sheet.data[cellId];
        }
      });

      return cloned;
    });

    setToastMessage("Duplicate records in range successfully cleared.");
  };

  // Split cell contents in column by custom delimiter (Text to columns)
  const handleTextToColumns = (delimiter: string) => {
    if (!activeSheet || !selectedCellId) {
      setToastMessage("Please select cell(s) to apply Text to Columns.");
      return;
    }

    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;

      let targetRows: number[] = [];
      let startColIdx = 0;

      if (activeRange) {
        const startParsed = parseCellId(activeRange.start);
        const endParsed = parseCellId(activeRange.end);
        if (startParsed && endParsed) {
          startColIdx = startParsed.colIndex;
          const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
          const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);
          for (let r = minRow; r <= maxRow; r++) {
            targetRows.push(r);
          }
        }
      } else {
        const parsed = parseCellId(selectedCellId);
        if (parsed) {
          startColIdx = parsed.colIndex;
          targetRows.push(parsed.rowIndex);
        }
      }

      let modifiedCount = 0;
      targetRows.forEach(r => {
        const sourceCellId = indexToColName(startColIdx) + r;
        const val = sheet.data[sourceCellId] || "";
        if (!val || val.startsWith("=")) return;

        const actualDelimiter = delimiter === "\\t" ? "\t" : delimiter;
        const parts = val.split(actualDelimiter);
        if (parts.length > 1) {
          modifiedCount++;
          parts.forEach((part, index) => {
            const targetColIdx = startColIdx + index;
            const targetCellId = indexToColName(targetColIdx) + r;
            sheet.data[targetCellId] = part.trim();
          });
        }
      });

      if (modifiedCount === 0) {
        return prev;
      }

      return cloned;
    });

    setToastMessage("Text successfully split across columns.");
  };

  // Column width resize state save handlers
  const handleResizeColumn = (col: string, width: number) => {
    setWorkbook(prev => {
      const cloned = { ...prev };
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (sheet) {
        sheet.columnWidths = {
          ...sheet.columnWidths,
          [col]: width
        };
      }
      return cloned;
    });
  };

  const handleInsertRow = (rowNum: number) => {
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      
      const newData: SheetData = {};
      const newStyles: SheetStyles = {};
      
      Object.entries(sheet.data).forEach(([cellId, val]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          if (r >= rowNum) {
            newData[col + (r + 1)] = val;
          } else {
            newData[cellId] = val;
          }
        }
      });
      
      Object.entries(sheet.styles).forEach(([cellId, style]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          if (r >= rowNum) {
            newStyles[col + (r + 1)] = style;
          } else {
            newStyles[cellId] = style;
          }
        }
      });
      
      sheet.data = newData;
      sheet.styles = newStyles;
      return cloned;
    });
    setToastMessage(`Inserted new row above row ${rowNum}`);
  };

  const handleDeleteRow = (rowNum: number) => {
    pushToHistory();
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      
      const newData: SheetData = {};
      const newStyles: SheetStyles = {};
      
      Object.entries(sheet.data).forEach(([cellId, val]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          if (r < rowNum) {
            newData[cellId] = val;
          } else if (r > rowNum) {
            newData[col + (r - 1)] = val;
          }
        }
      });
      
      Object.entries(sheet.styles).forEach(([cellId, style]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          if (r < rowNum) {
            newStyles[cellId] = style;
          } else if (r > rowNum) {
            newStyles[col + (r - 1)] = style;
          }
        }
      });
      
      sheet.data = newData;
      sheet.styles = newStyles;
      return cloned;
    });
    setToastMessage(`Deleted row ${rowNum}`);
  };

  const handleInsertColumn = (colName: string) => {
    pushToHistory();
    const targetColIdx = colNameToIndex(colName);
    
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      
      const newData: SheetData = {};
      const newStyles: SheetStyles = {};
      
      Object.entries(sheet.data).forEach(([cellId, val]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          const colIdx = parsed.colIndex;
          
          if (colIdx >= targetColIdx) {
            newData[indexToColName(colIdx + 1) + r] = val;
          } else {
            newData[cellId] = val;
          }
        }
      });
      
      Object.entries(sheet.styles).forEach(([cellId, style]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          const colIdx = parsed.colIndex;
          
          if (colIdx >= targetColIdx) {
            newStyles[indexToColName(colIdx + 1) + r] = style;
          } else {
            newStyles[cellId] = style;
          }
        }
      });
      
      sheet.data = newData;
      sheet.styles = newStyles;
      return cloned;
    });
    setToastMessage(`Inserted new column left of column ${colName}`);
  };

  const handleDeleteColumn = (colName: string) => {
    pushToHistory();
    const targetColIdx = colNameToIndex(colName);
    
    setWorkbook(prev => {
      const cloned = JSON.parse(JSON.stringify(prev)) as WorkbookState;
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return prev;
      
      const newData: SheetData = {};
      const newStyles: SheetStyles = {};
      
      Object.entries(sheet.data).forEach(([cellId, val]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          const colIdx = parsed.colIndex;
          
          if (colIdx < targetColIdx) {
            newData[cellId] = val;
          } else if (colIdx > targetColIdx) {
            newData[indexToColName(colIdx - 1) + r] = val;
          }
        }
      });
      
      Object.entries(sheet.styles).forEach(([cellId, style]) => {
        const parsed = parseCellId(cellId);
        if (parsed) {
          const col = parsed.col;
          const r = parsed.rowIndex;
          const colIdx = parsed.colIndex;
          
          if (colIdx < targetColIdx) {
            newStyles[cellId] = style;
          } else if (colIdx > targetColIdx) {
            newStyles[indexToColName(colIdx - 1) + r] = style;
          }
        }
      });
      
      sheet.data = newData;
      sheet.styles = newStyles;
      return cloned;
    });
    setToastMessage(`Deleted column ${colName}`);
  };
  
  // Clear drawings on current active sheet
  const handleClearDrawings = () => {
    const sheetId = workbook.activeSheetId;
    if (!sheetId) return;
    setDrawLines(prev => ({
      ...prev,
      [sheetId]: []
    }));
    setToastMessage("Cleared all drawings on current sheet");
  };

  // Insert a new chart or sticky note widget onto the sheet
  const handleInsertFloatingWidget = (type: 'chart' | 'note' | 'shape' | 'slicer' | 'pivot' | 'wordart', extra?: any) => {
    const sheetId = workbook.activeSheetId;
    if (!sheetId) return;

    let rangeStr = extra?.range;
    if (!rangeStr && activeRange) {
      rangeStr = `${activeRange.start}:${activeRange.end}`;
    }
    if (!rangeStr) {
      rangeStr = type === 'slicer' ? 'A' : 'A1:B4';
    }

    const newWidget: FloatingWidget = {
      id: `widget_${Date.now()}`,
      type,
      x: 160 + (floatingWidgets.length * 20) % 300,
      y: 120 + (floatingWidgets.length * 20) % 200,
      sheetId,
      title: extra?.title || (
        type === 'chart' ? 'Floating Chart' :
        type === 'shape' ? 'Custom Shape' :
        type === 'slicer' ? `Slicer ${rangeStr}` :
        type === 'pivot' ? 'Pivot Table' :
        type === 'wordart' ? 'WordArt Text' : 'Sticky Note'
      ),
      content: extra?.content || (type === 'note' ? 'Write some reminders or callouts here...' : ''),
      range: rangeStr,
      chartType: extra?.chartType || 'bar',
      shapeType: extra?.shapeType || 'rect',
      color: extra?.color,
      pivotRows: extra?.pivotRows || 'A',
      pivotCols: extra?.pivotCols || 'B',
      pivotVals: extra?.pivotVals || 'B',
      pivotOp: extra?.pivotOp || 'sum'
    };

    setFloatingWidgets(prev => [...prev, newWidget]);
    setToastMessage(`Inserted floating ${type}`);
  };

  // Execute JS macros for automated actions
  const handleRunMacroScript = (scriptText: string) => {
    try {
      const cloned = JSON.parse(JSON.stringify(workbook));
      const sheet = cloned.sheets[cloned.activeSheetId];
      if (!sheet) return;

      const context = {
        setCell: (cellId: string, val: any) => {
          sheet.data[cellId] = String(val);
        },
        getCell: (cellId: string) => {
          return sheet.data[cellId] || "";
        },
        setBold: (cellId: string, bold: boolean) => {
          if (!sheet.styles[cellId]) sheet.styles[cellId] = {};
          sheet.styles[cellId].bold = bold;
        },
        setItalic: (cellId: string, italic: boolean) => {
          if (!sheet.styles[cellId]) sheet.styles[cellId] = {};
          sheet.styles[cellId].italic = italic;
        },
        setBgColor: (cellId: string, color: string) => {
          if (!sheet.styles[cellId]) sheet.styles[cellId] = {};
          sheet.styles[cellId].bg = color;
        },
        setTextColor: (cellId: string, color: string) => {
          if (!sheet.styles[cellId]) sheet.styles[cellId] = {};
          sheet.styles[cellId].color = color;
        }
      };

      const runFn = new Function('ctx', `
        with(ctx) {
          ${scriptText}
        }
      `);

      runFn(context);

      setUndoStack(prev => [...prev, workbook]);
      setRedoStack([]);
      setWorkbook(cloned);
      setToastMessage("Macro executed successfully!");
    } catch (err: any) {
      console.error(err);
      setToastMessage(`Automation error: ${err.message}`);
    }
  };

  // Simple cell formula template insertion
  const handleInsertFormulaValue = (formulaTemplate: string) => {
    if (!selectedCellId) return;
    handleCellChanged(selectedCellId, formulaTemplate);
  };

  // Style of selected highlighted reference cell
  const currentCellStyle = useMemo(() => {
    if (!selectedCellId || !activeSheet || !activeSheet.styles) return {};
    return activeSheet.styles[selectedCellId] || {};
  }, [selectedCellId, activeSheet]);

  return (
    <div className="flex flex-col h-screen text-slate-800 bg-slate-100 font-sans">
      {/* Top Application Header bar */}
      <header className={`flex items-center justify-between px-6 py-2.5 ${themeStyles.bgHeader} text-white shadow-md transition-colors duration-200`}>
        <div className="flex items-center gap-2.5">
          <div className="p-1.5 bg-white/10 rounded-lg">
            <FileSpreadsheet className={`w-5 h-5 ${themeStyles.textBadge}`} />
          </div>
          <div className="flex flex-col">
            <h1 className="text-[14px] font-bold tracking-tight">Spreadsheet Pro</h1>
            <span className={`text-[10px] ${themeStyles.textSub} font-medium`}>Professional Formula & Execution Canvas</span>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <div className="hidden md:flex items-center gap-1.5 bg-white/10 rounded-full px-3 py-1 text-[11px] font-medium">
            <Info className={`w-3.5 h-3.5 ${themeStyles.textBadge}`} />
            <span className={themeStyles.textSub}>Formulas: SUM, AVERAGE, VLOOKUP, IF, DAYS, DATE, DATEDIF supported</span>
          </div>
          <button 
            type="button"
            onClick={() => setIsHelpOpen(true)}
            className={`flex items-center gap-1 text-xs text-white ${themeStyles.bgHeaderAccent} rounded-md px-2.5 py-1.5 transition-all cursor-pointer`}
          >
            <HelpCircle className="w-3.5 h-3.5" />
            <span>Help Tips</span>
          </button>
        </div>
      </header>

      {/* Dynamic Operational Toolbar */}
      <Toolbar 
        activeSheetName={activeSheet?.name || "Ledger"}
        onAddSheet={handleAddSheet}
        onRenameSheet={handleRenameSheet}
        onDeleteSheet={handleDeleteSheet}
        canUndo={undoStack.length > 0}
        canRedo={redoStack.length > 0}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onExportCSV={handleExportCSV}
        selectedCellStyle={currentCellStyle}
        onUpdateStyle={handleUpdateStyle}
        onSortActiveCol={handleSortActiveCol}
        onFindReplace={handleFindReplace}
        selectedCellId={selectedCellId}
        onAutoFillValue={handleAutoFillSelection}
        hasRangeSelected={activeRange !== null}
        onOpenImport={() => setIsImportOpen(true)}
        onInsertRow={handleInsertRow}
        onDeleteRow={handleDeleteRow}
        onInsertColumn={handleInsertColumn}
        onDeleteColumn={handleDeleteColumn}
        
        // Custom workbook customize layout and drawing props
        themeColor={themeColor}
        onThemeColorChange={setThemeColor}
        hideGridlines={hideGridlines}
        onHideGridlinesChange={setHideGridlines}
        drawModeEnabled={drawModeEnabled}
        onDrawModeEnabledChange={setDrawModeEnabled}
        drawColor={drawColor}
        onDrawColorChange={setDrawColor}
        drawThickness={drawThickness}
        onDrawThicknessChange={setDrawThickness}
        onClearDrawings={handleClearDrawings}
        onInsertFloatingWidget={handleInsertFloatingWidget}
        activeRange={activeRange}
        onRunMacroScript={handleRunMacroScript}
        onInsertFormulaValue={handleInsertFormulaValue}
        showFormulasActive={showFormulasActive}
        onShowFormulasActiveChange={setShowFormulasActive}

        onCopy={handleCopy}
        onCut={handleCut}
        onPaste={handlePaste}
        onPasteSpecial={handlePasteSpecial}
        hasClipboard={clipboard !== null}
        isFormatPainterActive={formatPainterStyle !== null}
        onToggleFormatPainter={handleToggleFormatPainter}
        onDuplicateSheet={handleDuplicateSheet}
        onCopyWorkbookBackup={handleCopyWorkbookBackup}
        onPasteWorkbookBackup={handlePasteWorkbookBackup}
        onRemoveDuplicates={handleRemoveDuplicates}
        onTextToColumns={handleTextToColumns}
        onCopySheetCSVText={handleCopySheetCSVText}
      />

      {/* Dynamic reactive formula editing toolbar header block */}
      <FormulaBar 
        selectedCellId={selectedCellId}
        cellNameInput={cellNameInput}
        onCellNameInputChange={setCellNameInput}
        onCellNameInputSubmit={handleCellNameInputSubmit}
        value={selectedCellFormula}
        onChange={(val) => handleCellChanged(selectedCellId!, val)}
      />

      {/* Main Table Spreadsheet scrollable Grid workspace */}
      <Grid 
        sheetData={activeSheet?.data || {}}
        evaluatedValues={evaluatedValues}
        sheetStyles={activeSheet?.styles || {}}
        selectedCellId={selectedCellId}
        activeRange={activeRange}
        columnWidths={activeSheet?.columnWidths || {}}
        rowHeights={activeSheet?.rowHeights || {}}
        hiddenColumns={activeSheet?.hiddenColumns || {}}
        hiddenRows={activeSheet?.hiddenRows || {}}
        onSelectCell={handleSelectCell}
        onStartRangeSelection={handleStartRangeSelection}
        onExpandRangeSelection={handleExpandRangeSelection}
        onEndRangeSelection={handleEndRangeSelection}
        onCellChanged={handleCellChanged}
        onResizeColumn={handleResizeColumn}
        onResizeRow={(row, height) => {
          setWorkbook(prev => {
            const cloned = { ...prev };
            const sheet = cloned.sheets[cloned.activeSheetId];
            if (sheet) {
              sheet.rowHeights = { ...sheet.rowHeights, [row]: height };
            }
            return cloned;
          });
        }}
        onHideColumn={handleHideColumn}
        onHideRow={handleHideRow}
        onUnhideAllColumns={handleUnhideAllColumns}
        onUnhideAllRows={handleUnhideAllRows}
        onCopy={handleCopy}
        onCut={handleCut}
        onPaste={handlePaste}
        hasClipboard={clipboard !== null}
        onInsertRow={handleInsertRow}
        onDeleteRow={handleDeleteRow}
        onInsertColumn={handleInsertColumn}
        onDeleteColumn={handleDeleteColumn}
        onAutoFillValue={handleAutoFillSelection}
        
        // Custom workbook customize layout and drawing props
        themeColor={themeColor}
        hideGridlines={hideGridlines}
        drawModeEnabled={drawModeEnabled}
        drawColor={drawColor}
        drawThickness={drawThickness}
        drawLines={drawLines}
        onDrawLinesChange={setDrawLines}
        floatingWidgets={floatingWidgets}
        onFloatingWidgetsChange={setFloatingWidgets}
        activeSheetId={workbook.activeSheetId}
        showFormulasActive={showFormulasActive}
      />

      {/* Tab bar sheets manager controls section */}
      <Tabs 
        sheets={workbook.sheets}
        sheetIds={workbook.sheetIds}
        activeSheetId={workbook.activeSheetId}
        onSelectSheet={(id) => { setWorkbook(prev => ({ ...prev, activeSheetId: id })); setSelectedCellId("A1"); setActiveRange(null); }}
        onAddSheet={handleAddSheet}
      />

      {/* File Import modal dialog */}
      <ImportModal 
        isOpen={isImportOpen}
        onClose={() => setIsImportOpen(false)}
        onImportData={handleImportData}
      />

      {/* Toast Notification Banner */}
      {toastMessage && (
        <div className="fixed bottom-14 right-5 z-50 flex items-center gap-3 bg-slate-900 text-white px-4 py-3 rounded-lg shadow-2xl animate-fade-in text-xs border border-slate-800">
          <AlertCircle className="w-4 h-4 text-emerald-400 flex-shrink-0" />
          <span>{toastMessage}</span>
          <button 
            onClick={() => setToastMessage(null)} 
            className="p-0.5 hover:bg-slate-800 rounded text-slate-400 hover:text-white cursor-pointer"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Help Tips Dialog Modal */}
      {isHelpOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs select-none">
          <div className="w-full max-w-md bg-white rounded-xl shadow-2xl border border-slate-200 p-6 flex flex-col gap-4 text-left">
            <div className="flex items-center justify-between border-b pb-3">
              <div className="flex items-center gap-2">
                <HelpCircle className="w-5 h-5 text-emerald-600" />
                <h3 className="text-sm font-bold text-slate-800">Spreadsheet Pro Guide</h3>
              </div>
              <button 
                onClick={() => setIsHelpOpen(false)} 
                className="p-1 hover:bg-slate-100 rounded-md text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            
            <div className="space-y-3.5 text-xs text-slate-600">
              <div className="flex items-start gap-2.5">
                <div className="w-5 h-5 rounded-full bg-emerald-50 text-emerald-700 font-bold flex items-center justify-center flex-shrink-0 text-[10px]">1</div>
                <div>
                  <p className="font-bold text-slate-800">Navigation shortcuts</p>
                  <p className="text-slate-500 mt-0.5">Use Arrow Keys, Tab, and Enter to move across the grid selection canvas, which updates details instantly.</p>
                </div>
              </div>
              <div className="flex items-start gap-2.5">
                <div className="w-5 h-5 rounded-full bg-emerald-50 text-emerald-700 font-bold flex items-center justify-center flex-shrink-0 text-[10px]">2</div>
                <div>
                  <p className="font-bold text-slate-800">Inline evaluation entry</p>
                  <p className="text-slate-500 mt-0.5">Double click or start typing directly when selection indicator highlights target grid cell coordinates.</p>
                </div>
              </div>
              <div className="flex items-start gap-2.5">
                <div className="w-5 h-5 rounded-full bg-emerald-50 text-emerald-700 font-bold flex items-center justify-center flex-shrink-0 text-[10px]">3</div>
                <div>
                  <p className="font-bold text-slate-800">Formulation Syntax</p>
                  <p className="text-slate-500 mt-0.5">Prefix with <code className="bg-slate-100 px-1 py-0.5 font-mono text-[10px] rounded text-emerald-700">=</code> or <code className="bg-slate-100 px-1 py-0.5 font-mono text-[10px] rounded">+</code> to execute formulas (e.g. <code className="bg-slate-100 px-1 py-0.5 font-mono text-[10px] rounded">=SUM(B4:B9)</code>, <code className="bg-slate-100 px-1 py-0.5 font-mono text-[10px] rounded">+AVERAGE(C5:E5)</code>).</p>
                </div>
              </div>
              <div className="flex items-start gap-2.5">
                <div className="w-5 h-5 rounded-full bg-emerald-50 text-emerald-700 font-bold flex items-center justify-center flex-shrink-0 text-[10px]">4</div>
                <div>
                  <p className="font-bold text-slate-800">Clipboard & Custom Actions</p>
                  <p className="text-slate-500 mt-0.5">Press <kbd className="bg-slate-100 px-1 py-0.5 rounded border border-slate-200">Ctrl+C</kbd> (copy), <kbd className="bg-slate-100 px-1 py-0.5 rounded border border-slate-200">Ctrl+X</kbd> (cut), or <kbd className="bg-slate-100 px-1 py-0.5 rounded border border-slate-200">Ctrl+V</kbd> (paste) to manipulate cells. Right-click column/row headers to hide columns or rows!</p>
                </div>
              </div>
            </div>

            <div className="flex justify-end pt-2 border-t">
              <button 
                onClick={() => setIsHelpOpen(false)}
                className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg text-xs shadow-sm cursor-pointer"
              >
                Got it, close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
