import React, { useState, useRef, useEffect } from 'react';
import { colNameToIndex, indexToColName, parseCellId } from '../utils/formulaEvaluator';
import { SheetData, SheetStyles, CellStyle, SelectionRange } from '../types';
import { Trash } from 'lucide-react';

interface GridProps {
  sheetData: SheetData;
  evaluatedValues: { [cellId: string]: string };
  sheetStyles: SheetStyles;
  selectedCellId: string | null;
  activeRange: SelectionRange | null;
  columnWidths: { [col: string]: number };
  rowHeights: { [row: number]: number };
  hiddenColumns: { [col: string]: boolean };
  hiddenRows: { [row: number]: boolean };
  onSelectCell: (cellId: string) => void;
  onStartRangeSelection: (cellId: string) => void;
  onExpandRangeSelection: (cellId: string) => void;
  onEndRangeSelection: () => void;
  onCellChanged: (cellId: string, value: string) => void;
  onResizeColumn: (col: string, width: number) => void;
  onResizeRow: (row: number, height: number) => void;
  onHideColumn: (col: string) => void;
  onHideRow: (row: number) => void;
  onUnhideAllColumns: () => void;
  onUnhideAllRows: () => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  hasClipboard: boolean;
  onInsertRow: (row: number) => void;
  onDeleteRow: (row: number) => void;
  onInsertColumn: (col: string) => void;
  onDeleteColumn: (col: string) => void;
  onAutoFillValue: () => void;

  // workbook customization & graphics states
  themeColor: 'green' | 'blue' | 'purple' | 'rose' | 'amber' | 'slate';
  hideGridlines: boolean;
  drawModeEnabled: boolean;
  drawColor: string;
  drawThickness: number;
  drawLines: { [sheetId: string]: Array<{ points: number[]; color: string; width: number }> };
  onDrawLinesChange: (lines: { [sheetId: string]: Array<{ points: number[]; color: string; width: number }> }) => void;
  floatingWidgets: Array<{
    id: string;
    type: 'chart' | 'note' | 'shape' | 'slicer' | 'pivot' | 'text' | 'wordart';
    x: number;
    y: number;
    sheetId: string;
    title: string;
    content?: string;
    range?: string;
    chartType?: 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'map' | 'pivot';
    shapeType?: 'rect' | 'circle' | 'arrow';
    color?: string;
    pivotRows?: string;
    pivotCols?: string;
    pivotVals?: string;
    pivotOp?: 'sum' | 'avg' | 'count' | 'max' | 'min';
  }>;
  onFloatingWidgetsChange: (widgets: Array<{
    id: string;
    type: 'chart' | 'note' | 'shape' | 'slicer' | 'pivot' | 'text' | 'wordart';
    x: number;
    y: number;
    sheetId: string;
    title: string;
    content?: string;
    range?: string;
    chartType?: 'bar' | 'line' | 'area' | 'pie' | 'scatter' | 'map' | 'pivot';
    shapeType?: 'rect' | 'circle' | 'arrow';
    color?: string;
    pivotRows?: string;
    pivotCols?: string;
    pivotVals?: string;
    pivotOp?: 'sum' | 'avg' | 'count' | 'max' | 'min';
  }>) => void;
  activeSheetId: string;
  showFormulasActive?: boolean;
}

const COL_COUNT = 1000; // 1000 columns
const ROW_COUNT = 1000; // 1000 rows

export const Grid: React.FC<GridProps> = ({
  sheetData,
  evaluatedValues,
  sheetStyles,
  selectedCellId,
  activeRange,
  columnWidths,
  rowHeights,
  hiddenColumns,
  hiddenRows,
  onSelectCell,
  onStartRangeSelection,
  onExpandRangeSelection,
  onEndRangeSelection,
  onCellChanged,
  onResizeColumn,
  onResizeRow,
  onHideColumn,
  onHideRow,
  onUnhideAllColumns,
  onUnhideAllRows,
  onCopy,
  onCut,
  onPaste,
  hasClipboard,
  onInsertRow,
  onDeleteRow,
  onInsertColumn,
  onDeleteColumn,
  onAutoFillValue,

  themeColor,
  hideGridlines,
  drawModeEnabled,
  drawColor,
  drawThickness,
  drawLines,
  onDrawLinesChange,
  floatingWidgets,
  onFloatingWidgetsChange,
  activeSheetId,
  showFormulasActive
}) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editingValue, setEditingValue] = useState("");
  const gridContainerRef = useRef<HTMLDivElement>(null);
  const activeInputRef = useRef<HTMLInputElement>(null);

  // Keep tracking mouse drags for range selection
  const [isMouseDown, setIsMouseDown] = useState(false);

  // Viewport virtual scrolling states
  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);
  const [containerHeight, setContainerHeight] = useState(600);
  const [containerWidth, setContainerWidth] = useState(1000);

  // Autocomplete Suggestions State
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [contextMenu, setContextMenu] = useState<{
    x: number;
    y: number;
    type: 'cell' | 'column' | 'row';
    target: string | number;
  } | null>(null);

  // Drawing Canvas and Floating Widget helpers
  const [currentLine, setCurrentLine] = useState<{ points: number[]; color: string; width: number } | null>(null);

  const activeSheetLines = drawLines[activeSheetId] || [];
  const sheetWidgets = floatingWidgets.filter(w => w.sheetId === activeSheetId);

  const handleDrawMouseDown = (e: React.MouseEvent<SVGSVGElement>) => {
    if (!drawModeEnabled) return;
    e.preventDefault();
    const rect = e.currentTarget.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    
    setCurrentLine({
      points: [x, y],
      color: drawColor,
      width: drawThickness
    });
  };

  const handleDrawMouseMove = (e: React.MouseEvent<SVGSVGElement>) => {
    if (!drawModeEnabled || !currentLine) return;
    e.preventDefault();
    const rect = e.currentTarget.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    setCurrentLine(prev => prev ? {
      ...prev,
      points: [...prev.points, x, y]
    } : null);
  };

  const handleDrawMouseUp = () => {
    if (!currentLine) return;
    const updatedLines = { ...drawLines };
    if (!updatedLines[activeSheetId]) {
      updatedLines[activeSheetId] = [];
    }
    updatedLines[activeSheetId] = [...updatedLines[activeSheetId], currentLine];
    onDrawLinesChange(updatedLines);
    setCurrentLine(null);
  };

  // Drag floating widgets
  const handleWidgetDragStart = (e: React.MouseEvent, widgetId: string) => {
    e.preventDefault();
    const widget = floatingWidgets.find(x => x.id === widgetId);
    if (!widget) return;
    const startX = e.clientX - widget.x;
    const startY = e.clientY - widget.y;

    const handleDragMove = (moveEvt: MouseEvent) => {
      const nextX = Math.max(10, moveEvt.clientX - startX);
      const nextY = Math.max(10, moveEvt.clientY - startY);
      onFloatingWidgetsChange(
        floatingWidgets.map(x => x.id === widgetId ? { ...x, x: nextX, y: nextY } : x)
      );
    };

    const handleDragEnd = () => {
      document.removeEventListener('mousemove', handleDragMove);
      document.removeEventListener('mouseup', handleDragEnd);
    };

    document.addEventListener('mousemove', handleDragMove);
    document.addEventListener('mouseup', handleDragEnd);
  };

  const handleDeleteWidget = (widgetId: string) => {
    onFloatingWidgetsChange(floatingWidgets.filter(x => x.id !== widgetId));
  };

  const handleWidgetTextChange = (widgetId: string, text: string) => {
    onFloatingWidgetsChange(
      floatingWidgets.map(x => x.id === widgetId ? { ...x, content: text } : x)
    );
  };

  // Dynamic range parsing for charts
  const parseRangeDataForChart = (rangeStr: string) => {
    if (!rangeStr) return [];
    const parts = rangeStr.split(':');
    if (parts.length !== 2) return [];
    
    const start = parseCellId(parts[0]);
    const end = parseCellId(parts[1]);
    if (!start || !end) return [];

    const minCol = Math.min(start.colIndex, end.colIndex);
    const maxCol = Math.max(start.colIndex, end.colIndex);
    const minRow = Math.min(start.rowIndex, end.rowIndex);
    const maxRow = Math.max(start.rowIndex, end.rowIndex);

    const dataPoints: Array<{ label: string, value: number }> = [];

    for (let r = minRow; r <= maxRow; r++) {
      const labelCellId = indexToColName(minCol) + r;
      const valueCellId = indexToColName(maxCol) + r;

      const labelVal = evaluatedValues[labelCellId] || labelCellId;
      const rawVal = evaluatedValues[valueCellId] || "0";
      const parsedNum = parseFloat(rawVal);
      
      dataPoints.push({
        label: labelVal.replace(/"/g, ''),
        value: isNaN(parsedNum) ? 0 : parsedNum
      });
    }
    return dataPoints;
  };

  const formatCellValue = (val: string, format?: string): string => {
    if (!val || !format || format === 'general') return val;
    const num = Number(val);
    if (isNaN(num)) return val;
    switch (format) {
      case 'number':
        return num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      case 'currency':
        return num.toLocaleString(undefined, { style: 'currency', currency: 'USD' });
      case 'percent':
        return (num * 100).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '%';
      case 'scientific':
        return num.toExponential(2);
      default:
        return val;
    }
  };

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

  const inlineSuggestions = getSuggestions(editingValue);

  // Close context menu on outside click
  useEffect(() => {
    const closeMenu = () => setContextMenu(null);
    window.addEventListener('click', closeMenu);
    return () => window.removeEventListener('click', closeMenu);
  }, []);

  // Generate cols header A to ALL, skipping hidden ones
  const cols = React.useMemo(() => {
    return Array.from({ length: COL_COUNT }, (_, i) => indexToColName(i))
      .filter(c => !hiddenColumns[c]);
  }, [hiddenColumns]);

  // Generate rows header 1 to ROW_COUNT, skipping hidden ones
  const rows = React.useMemo(() => {
    return Array.from({ length: ROW_COUNT }, (_, i) => i + 1)
      .filter(r => !hiddenRows[r]);
  }, [hiddenRows]);

  // Filter rows based on active slicers!
  const filteredRows = React.useMemo(() => {
    let result = rows;
    const activeSlicers = floatingWidgets.filter(w => w.type === 'slicer' && w.sheetId === activeSheetId && w.content);
    if (activeSlicers.length > 0) {
      result = result.filter(r => {
        return activeSlicers.every(slicer => {
          const col = slicer.range; // column name like "A"
          if (!col) return true;
          const cellId = col + r;
          const val = (evaluatedValues[cellId] || "").trim();
          if (!slicer.content || slicer.content === "__ALL__" || slicer.content === "") return true;
          const selectedVals = slicer.content.split(';');
          return selectedVals.includes(val);
        });
      });
    }
    return result;
  }, [rows, floatingWidgets, activeSheetId, evaluatedValues]);

  // Track scroll container size dynamically
  useEffect(() => {
    const handleResize = () => {
      if (gridContainerRef.current) {
        setContainerHeight(gridContainerRef.current.clientHeight);
        setContainerWidth(gridContainerRef.current.clientWidth);
      }
    };
    handleResize();
    const timer = setTimeout(handleResize, 100);
    window.addEventListener('resize', handleResize);
    return () => {
      window.removeEventListener('resize', handleResize);
      clearTimeout(timer);
    };
  }, []);

  const handleScroll = (e: React.UIEvent<HTMLDivElement>) => {
    const target = e.currentTarget;
    setScrollTop(target.scrollTop);
    setScrollLeft(target.scrollLeft);
  };

  // Measure visible row range
  let startRowIdx = 0;
  let accumulatedHeight = 0;
  while (startRowIdx < filteredRows.length && accumulatedHeight + (rowHeights[filteredRows[startRowIdx]] || 28) < scrollTop) {
    accumulatedHeight += (rowHeights[filteredRows[startRowIdx]] || 28);
    startRowIdx++;
  }
  const bufferRows = 15;
  const visibleStartRowIdx = Math.max(0, startRowIdx - bufferRows);

  let endRowIdx = startRowIdx;
  let visibleHeight = 0;
  while (endRowIdx < filteredRows.length && visibleHeight < containerHeight) {
    visibleHeight += (rowHeights[filteredRows[endRowIdx]] || 28);
    endRowIdx++;
  }
  const visibleEndRowIdx = Math.min(filteredRows.length - 1, endRowIdx + bufferRows);

  // Measure visible col range
  let startColIdx = 0;
  let accumulatedWidth = 0;
  while (startColIdx < cols.length) {
    const colName = cols[startColIdx];
    const w = columnWidths[colName] || 100;
    if (accumulatedWidth + w >= scrollLeft) break;
    accumulatedWidth += w;
    startColIdx++;
  }
  const bufferCols = 6;
  const visibleStartColIdx = Math.max(0, startColIdx - bufferCols);

  let endColIdx = startColIdx;
  let visibleWidth = 0;
  while (endColIdx < cols.length && visibleWidth < containerWidth) {
    const colName = cols[endColIdx];
    const w = columnWidths[colName] || 100;
    visibleWidth += w;
    endColIdx++;
  }
  const visibleEndColIdx = Math.min(cols.length - 1, endColIdx + bufferCols);

  // Slice visible arrays for virtual rendering
  const visibleCols = cols.slice(visibleStartColIdx, visibleEndColIdx + 1);
  const visibleRows = filteredRows.slice(visibleStartRowIdx, visibleEndRowIdx + 1);

  // Hidden left columns width:
  let hiddenLeftWidth = 0;
  for (let i = 0; i < visibleStartColIdx; i++) {
    hiddenLeftWidth += (columnWidths[cols[i]] || 100);
  }

  // Hidden right columns width:
  let hiddenRightWidth = 0;
  for (let i = visibleEndColIdx + 1; i < cols.length; i++) {
    hiddenRightWidth += (columnWidths[cols[i]] || 100);
  }

  // Hidden top rows height:
  let hiddenTopHeight = 0;
  for (let i = 0; i < visibleStartRowIdx; i++) {
    hiddenTopHeight += (rowHeights[filteredRows[i]] || 28);
  }

  // Hidden bottom rows height:
  let hiddenBottomHeight = 0;
  for (let i = visibleEndRowIdx + 1; i < filteredRows.length; i++) {
    hiddenBottomHeight += (rowHeights[filteredRows[i]] || 28);
  }

  // Scroll selected cell into view to support perfect navigational key scroll jumps
  useEffect(() => {
    if (selectedCellId && gridContainerRef.current) {
      const parsed = parseCellId(selectedCellId);
      if (!parsed) return;
      
      const container = gridContainerRef.current;
      const colName = parsed.col;
      const rowNum = parsed.rowIndex;
      
      if (hiddenColumns[colName] || hiddenRows[rowNum]) return; // Skip if cell is hidden

      // Calculate selected cell's accumulated offsets from the starting boundary
      let cellLeft = 48; // Left header column width
      for (let i = 0; i < cols.length; i++) {
        if (cols[i] === colName) break;
        cellLeft += (columnWidths[cols[i]] || 100);
      }
      const cellWidth = columnWidths[colName] || 100;
      
      let cellTop = 32; // Top header row height
      for (let i = 0; i < filteredRows.length; i++) {
        if (filteredRows[i] === rowNum) break;
        cellTop += (rowHeights[filteredRows[i]] || 28);
      }
      const cellHeight = rowHeights[rowNum] || 28;
      
      // Check horizontal boundary containment
      const scrollLeftMin = cellLeft + cellWidth - container.clientWidth;
      const scrollLeftMax = cellLeft - 48;
      if (container.scrollLeft < scrollLeftMin) {
        container.scrollLeft = scrollLeftMin + 20;
      } else if (container.scrollLeft > scrollLeftMax) {
        container.scrollLeft = Math.max(0, scrollLeftMax);
      }
      
      // Check vertical boundary containment
      const scrollTopMin = cellTop + cellHeight - container.clientHeight;
      const scrollTopMax = cellTop - 32;
      if (container.scrollTop < scrollTopMin) {
        container.scrollTop = scrollTopMin + 20;
      } else if (container.scrollTop > scrollTopMax) {
        container.scrollTop = Math.max(0, scrollTopMax);
      }
    }
  }, [selectedCellId]);

  // When active cell selection changes, ensure edit mode is turned off, the new value is initialized, and focus is placed on the cell element.
  useEffect(() => {
    if (selectedCellId) {
      setIsEditing(false);
      setEditingValue(sheetData[selectedCellId] || "");
      const element = document.getElementById(`cell-${selectedCellId}`);
      if (element) {
        element.focus();
      }
    } else {
      setIsEditing(false);
    }
  }, [selectedCellId]);

  // Keep editingValue synced with sheetData updates if we are not actively in edit mode (e.g. formula evaluation runs)
  useEffect(() => {
    if (selectedCellId && !isEditing) {
      setEditingValue(sheetData[selectedCellId] || "");
    }
  }, [sheetData]);

  // Retain focus and set cursor location to the end when editing is initiated
  useEffect(() => {
    if (isEditing && activeInputRef.current) {
      const input = activeInputRef.current;
      input.focus();
      const len = input.value.length;
      input.setSelectionRange(len, len);
    }
  }, [isEditing]);

  // Check if a cell is inside the selected range
  const isCellInRange = (cellId: string): boolean => {
    if (cellId === selectedCellId) return true;
    if (!activeRange) return false;

    const cellParsed = parseCellId(cellId);
    const startParsed = parseCellId(activeRange.start);
    const endParsed = parseCellId(activeRange.end);

    if (!cellParsed || !startParsed || !endParsed) return false;

    const minCol = Math.min(startParsed.colIndex, endParsed.colIndex);
    const maxCol = Math.max(startParsed.colIndex, endParsed.colIndex);
    const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
    const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);

    return (
      cellParsed.colIndex >= minCol &&
      cellParsed.colIndex <= maxCol &&
      cellParsed.rowIndex >= minRow &&
      cellParsed.rowIndex <= maxRow
    );
  };

  // Safe navigation function using cell index coordinate jumps
  const moveCursor = (dCol: number, dRow: number) => {
    if (!selectedCellId) return;
    const parsed = parseCellId(selectedCellId);
    if (!parsed) return;

    // We navigate using non-hidden columns and rows!
    const currentColName = parsed.col;
    const currentRowNum = parsed.rowIndex;
    
    const currentColIdxInFiltered = cols.indexOf(currentColName);
    const currentRowIdxInFiltered = filteredRows.indexOf(currentRowNum);

    if (currentColIdxInFiltered === -1 || currentRowIdxInFiltered === -1) return;

    let targetColIdxInFiltered = currentColIdxInFiltered + dCol;
    let targetRowIdxInFiltered = currentRowIdxInFiltered + dRow;

    // Constrain boundaries
    if (targetColIdxInFiltered < 0) targetColIdxInFiltered = 0;
    if (targetColIdxInFiltered >= cols.length) targetColIdxInFiltered = cols.length - 1;
    if (targetRowIdxInFiltered < 0) targetRowIdxInFiltered = 0;
    if (targetRowIdxInFiltered >= filteredRows.length) targetRowIdxInFiltered = filteredRows.length - 1;

    const targetId = cols[targetColIdxInFiltered] + filteredRows[targetRowIdxInFiltered];
    onSelectCell(targetId);
  };

  // Keyboard navigation on highlighted cells
  const handleCellKeyDown = (e: React.KeyboardEvent<HTMLDivElement>, cellId: string) => {
    if (isEditing) {
      if (e.key === 'Enter') {
        onCellChanged(cellId, editingValue);
        setIsEditing(false);
        e.preventDefault();
        moveCursor(0, 1);
      } else if (e.key === 'Escape') {
        setIsEditing(false);
        e.preventDefault();
      } else if (e.key === 'Tab') {
        onCellChanged(cellId, editingValue);
        setIsEditing(false);
        e.preventDefault();
        moveCursor(e.shiftKey ? -1 : 1, 0);
      } else if (e.key === 'ArrowUp') {
        onCellChanged(cellId, editingValue);
        setIsEditing(false);
        e.preventDefault();
        moveCursor(0, -1);
      } else if (e.key === 'ArrowDown') {
        onCellChanged(cellId, editingValue);
        setIsEditing(false);
        e.preventDefault();
        moveCursor(0, 1);
      }
      return;
    }

    // Keyboard clipboard shortcuts and Ctrl+A selection
    if (e.ctrlKey || e.metaKey) {
      const key = e.key.toLowerCase();
      if (key === 'a') {
        e.preventDefault();
        onStartRangeSelection("A1");
        onExpandRangeSelection(`${cols[cols.length - 1]}${rows[rows.length - 1]}`);
        return;
      }
      if (key === 'c') {
        e.preventDefault();
        onCopy();
        return;
      }
      if (key === 'x') {
        e.preventDefault();
        onCut();
        return;
      }
      if (key === 'v') {
        e.preventDefault();
        onPaste();
        return;
      }
      if (key === 'd') {
        e.preventDefault();
        onAutoFillValue();
        return;
      }
    }

    // Normal navigational key listener
    switch (e.key) {
      case 'ArrowUp':
        moveCursor(0, -1);
        e.preventDefault();
        break;
      case 'ArrowDown':
        moveCursor(0, 1);
        e.preventDefault();
        break;
      case 'ArrowLeft':
        moveCursor(-1, 0);
        e.preventDefault();
        break;
      case 'ArrowRight':
        moveCursor(1, 0);
        e.preventDefault();
        break;
      case 'Tab':
        moveCursor(e.shiftKey ? -1 : 1, 0);
        e.preventDefault();
        break;
      case 'Enter':
        setIsEditing(true);
        e.preventDefault();
        break;
      case 'F2':
        setIsEditing(true);
        e.preventDefault();
        break;
      case 'Delete':
      case 'Backspace':
        onCellChanged(cellId, "");
        e.preventDefault();
        break;
      default:
        // Begin typing direct value-replacement if alphanumeric or symbol keystroke is typed while cell highlighted
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
          setIsEditing(true);
          setEditingValue(e.key);
          e.preventDefault();
        }
        break;
    }
  };

  // Column Drag Sizing handling variables
  const activeColResize = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

  const startColResize = (e: React.MouseEvent, col: string) => {
    e.preventDefault();
    e.stopPropagation();
    const startWidth = columnWidths[col] || 100;
    activeColResize.current = {
      col,
      startX: e.clientX,
      startWidth,
    };
    document.addEventListener('mousemove', handleColResizeMove);
    document.addEventListener('mouseup', handleColResizeEnd);
  };

  const handleColResizeMove = (e: MouseEvent) => {
    if (!activeColResize.current) return;
    const { col, startX, startWidth } = activeColResize.current;
    const diff = e.clientX - startX;
    const newWidth = Math.max(50, startWidth + diff);
    onResizeColumn(col, newWidth);
  };

  const handleColResizeEnd = () => {
    activeColResize.current = null;
    document.removeEventListener('mousemove', handleColResizeMove);
    document.removeEventListener('mouseup', handleColResizeEnd);
  };

  // Cell Mouse Interaction handlers
  const handleCellMouseDown = (e: React.MouseEvent, cellId: string) => {
    if (e.button !== 0) return; // Only process left click
    
    // Check if we are currently editing a formula
    if (isEditing && (editingValue.startsWith("=") || editingValue.startsWith("+"))) {
      e.preventDefault();
      e.stopPropagation();
      // We want to insert the clicked cellId into our editingValue at the cursor position
      if (activeInputRef.current) {
        const input = activeInputRef.current;
        const start = input.selectionStart ?? editingValue.length;
        const end = input.selectionEnd ?? editingValue.length;
        const newValue = editingValue.slice(0, start) + cellId + editingValue.slice(end);
        setEditingValue(newValue);
        onCellChanged(selectedCellId!, newValue);
        
        // Restore focus and put cursor after the inserted cellId
        setTimeout(() => {
          input.focus();
          const newCursorPos = start + cellId.length;
          input.setSelectionRange(newCursorPos, newCursorPos);
        }, 0);
      } else {
        const newValue = editingValue + cellId;
        setEditingValue(newValue);
        onCellChanged(selectedCellId!, newValue);
      }
      return;
    }

    // Click on checkbox cell to toggle!
    const rawVal = sheetData[cellId] || "";
    if (rawVal === "[ ]" || rawVal === "[✓]" || rawVal === "[]" || rawVal === "[x]") {
      onCellChanged(cellId, rawVal.includes("✓") || rawVal.includes("x") ? "[ ]" : "[✓]");
      onSelectCell(cellId);
      return;
    }

    setIsMouseDown(true);
    
    // Check if user is clicking directly on cell editor input to not intercept focus
    if (isEditing && selectedCellId === cellId) return;

    if (e.shiftKey && selectedCellId) {
      // Create selective range from active to clicked cell
      onExpandRangeSelection(cellId);
    } else {
      onSelectCell(cellId);
      onStartRangeSelection(cellId);
    }
  };

  const handleCellMouseEnter = (cellId: string) => {
    if (isMouseDown) {
      onExpandRangeSelection(cellId);
    }
  };

  const handleGlobalMouseUp = () => {
    setIsMouseDown(false);
    onEndRangeSelection();
  };

  useEffect(() => {
    window.addEventListener('mouseup', handleGlobalMouseUp);
    return () => {
      window.removeEventListener('mouseup', handleGlobalMouseUp);
    };
  }, [isMouseDown]);

  return (
    <div 
      ref={gridContainerRef}
      onScroll={handleScroll}
      className="flex-1 overflow-auto bg-slate-100 select-none border-b border-slate-200 relative"
    >
      <div className="relative inline-block min-w-full">
        <table className="border-separate border-spacing-0 table-fixed min-w-full">
        <thead>
          <tr className="h-8">
            {/* Top-Left frozen corner cell */}
            <th className="sticky top-0 left-0 z-30 bg-slate-200 border-r border-b border-slate-300 w-12 text-center text-xs font-bold text-slate-500"></th>
            
            {/* Left hidden spacer column */}
            {hiddenLeftWidth > 0 && (
              <th style={{ width: `${hiddenLeftWidth}px`, minWidth: `${hiddenLeftWidth}px` }} className="sticky top-0 z-20 h-8 bg-slate-100 border-r border-b border-slate-300"></th>
            )}

            {/* Dynamic letters and header resizing */}
            {visibleCols.map((colName) => {
              const colWidth = columnWidths[colName] || 100;
              return (
                <th 
                  key={colName}
                  style={{ width: `${colWidth}px`, minWidth: `${colWidth}px` }} 
                  onContextMenu={(e) => {
                    e.preventDefault();
                    setContextMenu({
                      x: e.clientX,
                      y: e.clientY,
                      type: 'column',
                      target: colName
                    });
                  }}
                  className="relative sticky top-0 z-20 h-8 bg-slate-100 hover:bg-slate-200 border-r border-b border-slate-300 text-center text-xs font-bold text-slate-600 transition-colors cursor-context-menu"
                >
                  <span className="inline-block py-1 px-2 select-none">{colName}</span>
                  {/* Resize handle handle overlay */}
                  <div 
                    onMouseDown={(e) => startColResize(e, colName)}
                    className="absolute top-0 right-0 h-full w-2 cursor-col-resize hover:bg-green-400/80 active:bg-green-500 transition-colors z-30"
                    title="Drag to resize column"
                  />
                </th>
              );
            })}

            {/* Right hidden spacer column */}
            {hiddenRightWidth > 0 && (
              <th style={{ width: `${hiddenRightWidth}px`, minWidth: `${hiddenRightWidth}px` }} className="sticky top-0 z-20 h-8 bg-slate-100 border-b border-slate-300"></th>
            )}
          </tr>
        </thead>

        <tbody>
          {/* Top hidden spacer row */}
          {hiddenTopHeight > 0 && (
            <tr style={{ height: `${hiddenTopHeight}px` }}>
              <td colSpan={visibleCols.length + 5}></td>
            </tr>
          )}

          {visibleRows.map((rowNum) => {
            const rowHeight = rowHeights[rowNum] || 28;
            return (
              <tr key={rowNum} style={{ height: `${rowHeight}px` }}>
                {/* Sticky Left row marker index cells */}
                <td 
                  onContextMenu={(e) => {
                    e.preventDefault();
                    setContextMenu({
                      x: e.clientX,
                      y: e.clientY,
                      type: 'row',
                      target: rowNum
                    });
                  }}
                  className="sticky left-0 z-10 w-12 bg-slate-100 hover:bg-slate-200 cursor-context-menu border-r border-b border-slate-350 text-center text-[11px] font-bold text-slate-500 font-mono select-none"
                >
                  {rowNum}
                </td>

                {/* Left hidden columns spacer */}
                {hiddenLeftWidth > 0 && (
                  <td style={{ width: `${hiddenLeftWidth}px` }} className="border-r border-b border-slate-200"></td>
                )}

                {/* Render spreadsheet inputs columns in this row */}
                {visibleCols.map((colName) => {
                  const cellId = colName + rowNum;
                  const rawVal = sheetData[cellId] || "";
                  const evalVal = showFormulasActive ? rawVal : (evaluatedValues[cellId] !== undefined ? evaluatedValues[cellId] : rawVal);
                  const isSelected = selectedCellId === cellId;
                  const isHighlighted = isCellInRange(cellId);
                  
                  // Retrieve custom element formatting styles
                  const style: CellStyle = sheetStyles[cellId] || {};
                  
                  let selectionBorderColor = '#16a34a'; // default green
                  let highlightBgColor = '#f0fdf4';
                  let highlightBorderColor = '#86efac';
                  
                  if (themeColor === 'blue') {
                    selectionBorderColor = '#2563eb';
                    highlightBgColor = '#eff6ff';
                    highlightBorderColor = '#93c5fd';
                  } else if (themeColor === 'purple') {
                    selectionBorderColor = '#9333ea';
                    highlightBgColor = '#f3e8ff';
                    highlightBorderColor = '#d8b4fe';
                  } else if (themeColor === 'rose') {
                    selectionBorderColor = '#e11d48';
                    highlightBgColor = '#fff1f2';
                    highlightBorderColor = '#fda4af';
                  } else if (themeColor === 'amber') {
                    selectionBorderColor = '#d97706';
                    highlightBgColor = '#fefbeb';
                    highlightBorderColor = '#fde047';
                  } else if (themeColor === 'slate') {
                    selectionBorderColor = '#475569';
                    highlightBgColor = '#f1f5f9';
                    highlightBorderColor = '#cbd5e1';
                  }
                  
                  const computedStyle: React.CSSProperties = {
                    fontWeight: style.bold ? 'bold' : 'normal',
                    fontStyle: style.italic ? 'italic' : 'normal',
                    textDecoration: style.underline ? 'underline' : 'none',
                    color: style.color || '#1e293b',
                    backgroundColor: style.bg || (isHighlighted ? highlightBgColor : '#ffffff'),
                    textAlign: style.align || 'left',
                    fontSize: style.fontSize ? `${style.fontSize}px` : '13px',
                    borderColor: isSelected ? selectionBorderColor : (isHighlighted ? highlightBorderColor : (hideGridlines ? 'transparent' : '#e2e8f0')),
                  };

                  if (style.border === 'all') {
                    computedStyle.border = '1px solid #64748b';
                  } else if (style.border === 'bottom') {
                    computedStyle.borderBottom = '2px solid #0f172a';
                  } else if (style.border === 'top') {
                    computedStyle.borderTop = '2px solid #0f172a';
                  } else if (style.border === 'thick') {
                    computedStyle.border = '2px solid #0f172a';
                  } else if (style.border === 'double-bottom') {
                    computedStyle.borderBottom = '4px double #0f172a';
                  }

                  return (
                    <td 
                      key={cellId}
                      id={`cell-${cellId}`}
                      tabIndex={0}
                      onKeyDown={(e) => handleCellKeyDown(e, cellId)}
                      onMouseDown={(e) => handleCellMouseDown(e, cellId)}
                      onMouseEnter={() => handleCellMouseEnter(cellId)}
                      onContextMenu={(e) => {
                        e.preventDefault();
                        onSelectCell(cellId);
                        setContextMenu({
                          x: e.clientX,
                          y: e.clientY,
                          type: 'cell',
                          target: cellId
                        });
                      }}
                      onDoubleClick={() => { setIsEditing(true); setEditingValue(rawVal); }}
                      style={computedStyle}
                      className={`border-r border-b px-1 py-1.5 align-middle outline-none relative cursor-cell transition-all group overflow-hidden truncate ${
                        isSelected 
                          ? 'ring-2 ring-emerald-500 ring-inset z-10 font-medium' 
                          : isHighlighted 
                            ? 'ring-1 ring-green-300 ring-inset z-5 bg-green-50/40' 
                            : 'hover:border-slate-400'
                      }`}
                    >
                      {/* Active inline text selector input */}
                      {isEditing && isSelected ? (
                        <div className="absolute inset-0 z-30 flex flex-col">
                          <input
                            ref={activeInputRef}
                            type="text"
                            value={editingValue}
                            onChange={(e) => {
                              setEditingValue(e.target.value);
                              setShowSuggestions(true);
                            }}
                            onFocus={() => setShowSuggestions(true)}
                            onBlur={() => setTimeout(() => {
                              onCellChanged(cellId, editingValue);
                              setIsEditing(false);
                              setShowSuggestions(false);
                            }, 250)}
                            className="w-full h-full text-xs font-mono px-1 py-0 bg-white border border-green-600 focus:outline-hidden text-slate-900"
                          />
                          {showSuggestions && inlineSuggestions.length > 0 && (
                            <div className="absolute left-0 top-full mt-0.5 z-50 min-w-[150px] bg-white border border-slate-250 rounded shadow-lg max-h-36 overflow-y-auto py-0.5">
                              {inlineSuggestions.map(sug => (
                                <button
                                  key={sug}
                                  type="button"
                                  onMouseDown={() => {
                                    const match = editingValue.match(/([A-Z_]+)$/i);
                                    if (match) {
                                      const idx = match.index ?? 0;
                                      const newVal = editingValue.slice(0, idx) + sug + "(";
                                      setEditingValue(newVal);
                                      onCellChanged(cellId, newVal);
                                    }
                                    setShowSuggestions(false);
                                  }}
                                  className="w-full text-left px-2.5 py-1 text-[11px] font-mono hover:bg-slate-100 text-slate-700 cursor-pointer flex items-center justify-between"
                                >
                                  <span>{sug}</span>
                                </button>
                              ))}
                            </div>
                          )}
                        </div>
                      ) : (
                        <div className="w-full h-full select-none truncate text-ellipsis overflow-hidden px-1 whitespace-pre">
                          {formatCellValue(evalVal, style.format)}
                        </div>
                      )}

                      {/* Cute Excel-like selection/fill handles */}
                      {isSelected && !isEditing && (
                        <div 
                          className="absolute bottom-0 right-0 w-2 h-2 bg-emerald-600 border border-white cursor-crosshair z-20 shadow-xs"
                          title="Drag selection"
                        />
                      )}

                      {/* Accent corner icon helper to represent existing spreadsheet state visually */}
                      {(rawVal.startsWith("=") || rawVal.startsWith("+")) && (
                        <div 
                          className="absolute right-0.5 top-0.5 w-1 h-1 bg-yellow-500 rounded-full" 
                          title={`Formula: ${rawVal}`}
                        />
                      )}
                    </td>
                  );
                })}

                {/* Right hidden columns spacer */}
                {hiddenRightWidth > 0 && (
                  <td style={{ width: `${hiddenRightWidth}px` }} className="border-b border-slate-250"></td>
                )}
              </tr>
            );
          })}

          {/* Bottom hidden spacer row */}
          {hiddenBottomHeight > 0 && (
            <tr style={{ height: `${hiddenBottomHeight}px` }}>
              <td colSpan={visibleCols.length + 5}></td>
            </tr>
          )}
        </tbody>
      </table>

      {/* Dynamic Drawing Layer */}
      {drawModeEnabled && (
        <svg
          onMouseDown={handleDrawMouseDown}
          onMouseMove={handleDrawMouseMove}
          onMouseUp={handleDrawMouseUp}
          onMouseLeave={handleDrawMouseUp}
          className="absolute top-0 left-0 w-full h-full cursor-crosshair z-10"
          style={{ minHeight: '100%', minWidth: '100%' }}
        >
          {activeSheetLines.map((line, idx) => (
            <polyline
              key={idx}
              points={line.points.join(',')}
              stroke={line.color}
              strokeWidth={line.width}
              fill="none"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          ))}
          {currentLine && (
            <polyline
              points={currentLine.points.join(',')}
              stroke={drawColor}
              strokeWidth={drawThickness}
              fill="none"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          )}
        </svg>
      )}

      {/* Dynamic Floating Widgets (sticky notes & charts) */}
      {sheetWidgets.map((w) => {
        return (
          <div
            key={w.id}
            style={{ left: `${w.x}px`, top: `${w.y}px` }}
            className={`absolute z-30 w-72 min-h-[220px] rounded-lg shadow-lg border p-3 flex flex-col bg-white border-slate-200 select-none animate-fade-in`}
          >
            {/* Header/Drag handle */}
            <div
              onMouseDown={(e) => handleWidgetDragStart(e, w.id)}
              className="flex items-center justify-between cursor-move bg-slate-50 border-b -mx-3 -mt-3 p-2 rounded-t-lg select-none"
            >
              <div className="flex items-center gap-1.5">
                <span className="text-[10px] uppercase font-bold text-slate-500 tracking-wider">
                  {w.type === 'chart' ? `📊 Chart (${w.chartType})` : 
                   w.type === 'shape' ? `🎨 Shape (${w.shapeType})` :
                   w.type === 'slicer' ? `🔍 Slicer (${w.range})` :
                   w.type === 'pivot' ? `🎛️ Pivot Table` :
                   w.type === 'wordart' ? `✨ WordArt` : `✍ Sticky Note`}
                </span>
                {w.range && (
                  <span className="text-[9px] font-mono bg-slate-200 text-slate-600 px-1 py-0.2 rounded font-bold">
                    {w.range}
                  </span>
                )}
              </div>
              <button
                onClick={() => handleDeleteWidget(w.id)}
                className="p-0.5 hover:bg-slate-200 text-slate-400 hover:text-red-500 rounded transition-colors"
                title="Close widget"
              >
                <Trash className="w-3.5 h-3.5" />
              </button>
            </div>

            {/* Widget Body */}
            <div className="flex-1 flex flex-col mt-3 overflow-hidden select-text">
              {w.type === 'note' && (
                <textarea
                  value={w.content || ""}
                  onChange={(e) => handleWidgetTextChange(w.id, e.target.value)}
                  className="w-full flex-1 text-xs border border-transparent hover:border-slate-100 focus:border-slate-200 focus:bg-amber-50/50 p-1.5 focus:outline-hidden font-medium text-slate-700 bg-transparent resize-none leading-relaxed"
                  placeholder="Type comments, notes, or explanations for this sheet..."
                />
              )}

              {w.type === 'shape' && (
                <div className="w-full flex-1 flex items-center justify-center min-h-[140px]">
                  {w.shapeType === 'rect' && (
                    <div className="w-full h-full min-h-[120px] flex items-center justify-center bg-indigo-500/80 border-2 border-indigo-600 rounded-md shadow-inner text-white text-xs font-bold p-3 text-center select-none">
                      {w.title}
                    </div>
                  )}
                  {w.shapeType === 'circle' && (
                    <div className="w-[120px] h-[120px] flex items-center justify-center bg-emerald-500/80 border-2 border-emerald-600 rounded-full shadow-inner text-white text-xs font-bold p-3 text-center select-none mx-auto">
                      {w.title}
                    </div>
                  )}
                  {w.shapeType === 'arrow' && (
                    <div className="w-full h-full min-h-[120px] flex flex-col items-center justify-center select-none">
                      <svg className="w-16 h-16 text-indigo-500 fill-current drop-shadow-sm rotate-90" viewBox="0 0 24 24">
                        <path d="M12 2L22 12H17V22H7V12H2L12 2Z" />
                      </svg>
                      <span className="text-[10px] text-slate-500 mt-1.5 font-bold uppercase tracking-wider">{w.title}</span>
                    </div>
                  )}
                </div>
              )}

              {w.type === 'wordart' && (
                <div className="w-full flex-1 flex items-center justify-center min-h-[140px] p-2">
                  <h1 style={{ color: w.color || '#4f46e5' }} className="text-2xl font-black tracking-wider text-center uppercase select-none drop-shadow-sm italic">
                    {w.content || w.title}
                  </h1>
                </div>
              )}

              {w.type === 'slicer' && (
                (() => {
                  const colLetter = w.range || "A";
                  const uniqueVals = React.useMemo(() => {
                    const vals = new Set<string>();
                    for (let r = 1; r <= 100; r++) {
                      const val = (evaluatedValues[colLetter + r] || "").trim();
                      if (val) vals.add(val);
                    }
                    return Array.from(vals).sort();
                  }, [colLetter, evaluatedValues]);

                  const selectedVals = w.content && w.content !== "__ALL__" ? w.content.split(';') : [];

                  const handleToggleVal = (v: string) => {
                    let nextVals;
                    if (selectedVals.includes(v)) {
                      nextVals = selectedVals.filter(x => x !== v);
                    } else {
                      nextVals = [...selectedVals, v];
                    }
                    const nextContent = nextVals.length === 0 ? "__ALL__" : nextVals.join(';');
                    onFloatingWidgetsChange(
                      floatingWidgets.map(x => x.id === w.id ? { ...x, content: nextContent } : x)
                    );
                  };

                  const handleClearSlicer = () => {
                    onFloatingWidgetsChange(
                      floatingWidgets.map(x => x.id === w.id ? { ...x, content: "__ALL__" } : x)
                    );
                  };

                  return (
                    <div className="flex-1 flex flex-col gap-2 overflow-hidden bg-slate-50 border border-slate-150 rounded p-2 select-none">
                      <div className="flex items-center justify-between text-[11px] font-bold text-slate-600 border-b pb-1">
                        <span>Filter Col {colLetter}</span>
                        <button onClick={handleClearSlicer} className="text-blue-600 hover:underline cursor-pointer text-[10px]">Clear</button>
                      </div>
                      <div className="flex-1 overflow-y-auto max-h-[140px] flex flex-col gap-1 pr-1">
                        {uniqueVals.length === 0 ? (
                          <span className="text-[10px] text-slate-400 italic">No values in Col {colLetter}</span>
                        ) : (
                          uniqueVals.map(v => {
                            const checked = selectedVals.length === 0 || selectedVals.includes(v);
                            return (
                              <label key={v} className="flex items-center gap-2 hover:bg-slate-100 p-1 rounded cursor-pointer text-xs select-none">
                                <input
                                  type="checkbox"
                                  checked={checked}
                                  onChange={() => handleToggleVal(v)}
                                  className="rounded text-green-600 focus:ring-green-500 w-3.5 h-3.5 border-slate-350 cursor-pointer"
                                />
                                <span className="truncate text-slate-700 font-medium">{v}</span>
                              </label>
                            );
                          })
                        )}
                      </div>
                    </div>
                  );
                })()
              )}

              {w.type === 'pivot' && (
                (() => {
                  const startCell = w.range?.split(':')[0] || "A1";
                  const endCell = w.range?.split(':')[1] || "A1";
                  const startParsed = parseCellId(startCell);
                  const endParsed = parseCellId(endCell);
                  if (!startParsed || !endParsed) return <div className="text-xs text-slate-400 italic">Invalid selection range</div>;

                  const rowCol = w.pivotRows || "A";
                  const valCol = w.pivotVals || "B";
                  const op = w.pivotOp || "sum";

                  const pivotData = React.useMemo(() => {
                    const groups: { [key: string]: number[] } = {};
                    const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
                    const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);

                    for (let r = minRow; r <= maxRow; r++) {
                      const groupKey = (evaluatedValues[rowCol + r] || "").trim();
                      const rawNum = parseFloat(evaluatedValues[valCol + r] || "0");
                      if (!groupKey) continue;
                      if (!groups[groupKey]) groups[groupKey] = [];
                      if (!isNaN(rawNum)) {
                        groups[groupKey].push(rawNum);
                      }
                    }

                    return Object.entries(groups).map(([key, nums]) => {
                      let aggregated = 0;
                      if (op === 'sum') {
                        aggregated = nums.reduce((a, b) => a + b, 0);
                      } else if (op === 'avg') {
                        aggregated = nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
                      } else if (op === 'count') {
                        aggregated = nums.length;
                      } else if (op === 'max') {
                        aggregated = nums.length ? Math.max(...nums) : 0;
                      } else if (op === 'min') {
                        aggregated = nums.length ? Math.min(...nums) : 0;
                      }
                      return { key, value: aggregated };
                    });
                  }, [rowCol, valCol, op, startParsed, endParsed, evaluatedValues]);

                  return (
                    <div className="flex-1 flex flex-col overflow-hidden select-none bg-slate-50 border border-slate-150 rounded p-1.5">
                      <div className="text-[10px] font-bold text-slate-500 uppercase tracking-wide border-b pb-1 mb-1.5 flex justify-between">
                        <span>Row: {rowCol} → Val: {valCol} ({op.toUpperCase()})</span>
                      </div>
                      <div className="flex-1 overflow-y-auto max-h-[140px]">
                        <table className="w-full text-left text-[11px] border-collapse">
                          <thead>
                            <tr className="border-b border-slate-200 bg-slate-100 text-slate-600">
                              <th className="p-1 font-bold">Group ({rowCol})</th>
                              <th className="p-1 font-bold text-right">{op.toUpperCase()} ({valCol})</th>
                            </tr>
                          </thead>
                          <tbody>
                            {pivotData.length === 0 ? (
                              <tr>
                                <td colSpan={2} className="p-2 text-center italic text-[10px] text-slate-400">No pivot data</td>
                              </tr>
                            ) : (
                              pivotData.map((row, i) => (
                                <tr key={i} className="border-b border-slate-100 hover:bg-slate-100">
                                  <td className="p-1 font-medium text-slate-700 truncate max-w-[120px]">{row.key}</td>
                                  <td className="p-1 text-right font-mono text-slate-800">{row.value.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  );
                })()
              )}

              {w.type === 'chart' && (
                /* Render dynamic SVG chart from range values! */
                (() => {
                  const chartData = parseRangeDataForChart(w.range || "");
                  if (chartData.length === 0) {
                    return (
                      <div className="flex-1 flex flex-col items-center justify-center text-slate-400 text-center text-xs p-4 border border-dashed border-slate-200 rounded-md">
                        <span>No numerical data found inside selection range.</span>
                      </div>
                    );
                  }
                  const maxVal = Math.max(...chartData.map(d => d.value), 1);
                  const svgH = 130;
                  const svgW = 260;
                  const pad = 24;

                  return (
                    <div className="flex-1 flex flex-col">
                      <svg className="w-full h-32 bg-slate-50/40 rounded-md border border-slate-100" viewBox={`0 0 ${svgW} ${svgH}`}>
                        {/* Y-axis helper grids */}
                        {[0, 0.25, 0.5, 0.75, 1].map((p, i) => {
                          const y = pad + p * (svgH - 2 * pad);
                          const valLabel = Math.round(maxVal * (1 - p));
                          return (
                            <g key={i} className="opacity-40">
                              <line x1={pad} y1={y} x2={svgW - pad} y2={y} stroke="#cbd5e1" strokeWidth="1" strokeDasharray="2,2" />
                              <text x={pad - 4} y={y + 3} textAnchor="end" fontSize="7" fill="#64748b" className="font-mono">{valLabel}</text>
                            </g>
                          );
                        })}

                        {/* Data visualizations */}
                        {w.chartType === 'bar' && (
                          chartData.map((d, i) => {
                            const step = (svgW - 2 * pad) / chartData.length;
                            const x = pad + i * step + step * 0.15;
                            const wBar = step * 0.7;
                            const hBar = (d.value / maxVal) * (svgH - 2 * pad);
                            const y = svgH - pad - hBar;

                            return (
                              <g key={i}>
                                <rect
                                  x={x}
                                  y={y}
                                  width={wBar}
                                  height={hBar}
                                  fill="url(#barGradient)"
                                  rx="2"
                                  className="hover:fill-indigo-600 transition-colors"
                                >
                                  <title>{d.label}: {d.value}</title>
                                </rect>
                                <text x={x + wBar / 2} y={svgH - pad + 8} textAnchor="middle" fontSize="6.5" fill="#475569" fontWeight="medium">
                                  {d.label.substring(0, 5)}
                                </text>
                              </g>
                            );
                          })
                        )}

                        {w.chartType === 'line' && (
                          (() => {
                            const step = (svgW - 2 * pad) / Math.max(chartData.length - 1, 1);
                            const pointsStr = chartData.map((d, i) => {
                              const x = pad + i * step;
                              const y = svgH - pad - (d.value / maxVal) * (svgH - 2 * pad);
                              return `${x},${y}`;
                            }).join(' ');

                            return (
                              <g>
                                {/* Line Path */}
                                <polyline points={pointsStr} fill="none" stroke="#6366f1" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
                                {/* Node Dots */}
                                {chartData.map((d, i) => {
                                  const x = pad + i * step;
                                  const y = svgH - pad - (d.value / maxVal) * (svgH - 2 * pad);
                                  return (
                                    <g key={i}>
                                      <circle cx={x} cy={y} r="3.5" fill="#ffffff" stroke="#6366f1" strokeWidth="2" className="hover:scale-125 transition-transform" />
                                      <text x={x} y={svgH - pad + 8} textAnchor="middle" fontSize="6.5" fill="#475569">
                                        {d.label.substring(0, 5)}
                                      </text>
                                    </g>
                                  );
                                })}
                              </g>
                            );
                          })()
                        )}

                        {w.chartType === 'area' && (
                          (() => {
                            const step = (svgW - 2 * pad) / Math.max(chartData.length - 1, 1);
                            const pointsStr = chartData.map((d, i) => {
                              const x = pad + i * step;
                              const y = svgH - pad - (d.value / maxVal) * (svgH - 2 * pad);
                              return `${x},${y}`;
                            }).join(' ');

                            const fillPath = `M ${pad},${svgH - pad} ` + chartData.map((d, i) => {
                              const x = pad + i * step;
                              const y = svgH - pad - (d.value / maxVal) * (svgH - 2 * pad);
                              return `L ${x},${y}`;
                            }).join(' ') + ` L ${pad + (chartData.length - 1) * step},${svgH - pad} Z`;

                            return (
                              <g>
                                {/* Area background */}
                                <path d={fillPath} fill="url(#areaFillGradient)" opacity="0.3" />
                                {/* Line over area */}
                                <polyline points={pointsStr} fill="none" stroke="#4f46e5" strokeWidth="2" />
                                {chartData.map((d, i) => {
                                  const x = pad + i * step;
                                  return (
                                    <text key={i} x={x} y={svgH - pad + 8} textAnchor="middle" fontSize="6.5" fill="#475569">
                                      {d.label.substring(0, 5)}
                                    </text>
                                  );
                                })}
                              </g>
                            );
                          })()
                        )}

                        {/* Gradients definitions inside SVG */}
                        <defs>
                          <linearGradient id="barGradient" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="#818cf8" />
                            <stop offset="100%" stopColor="#4f46e5" />
                          </linearGradient>
                          <linearGradient id="areaFillGradient" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="0%" stopColor="#6366f1" />
                            <stop offset="100%" stopColor="#e0e7ff" />
                          </linearGradient>
                        </defs>
                      </svg>
                    </div>
                  );
                })()
              )}
            </div>
          </div>
        );
      })}
      </div>

      {/* Floating Right Click Context Menu */}
      {contextMenu && (
        <div 
          style={{ top: `${contextMenu.y}px`, left: `${contextMenu.x}px` }}
          className="fixed z-50 bg-white border border-slate-250 rounded-lg shadow-xl py-1 min-w-[160px] text-xs text-slate-700 animate-fade-in"
        >
          {contextMenu.type === 'column' && (
            <>
              <button 
                type="button"
                onClick={() => { onInsertColumn(String(contextMenu.target)); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 font-medium text-green-700"
              >
                Insert Column Left
              </button>
              <button 
                type="button"
                onClick={() => { onDeleteColumn(String(contextMenu.target)); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 text-red-600"
              >
                Delete Column {contextMenu.target}
              </button>
              <button 
                type="button"
                onClick={() => { onHideColumn(String(contextMenu.target)); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 border-t border-slate-100"
              >
                Hide Column {contextMenu.target}
              </button>
              <button 
                type="button"
                onClick={() => { onUnhideAllColumns(); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2"
              >
                Unhide All Columns
              </button>
            </>
          )}
          {contextMenu.type === 'row' && (
            <>
              <button 
                type="button"
                onClick={() => { onInsertRow(Number(contextMenu.target)); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 font-medium text-green-700"
              >
                Insert Row Above
              </button>
              <button 
                type="button"
                onClick={() => { onDeleteRow(Number(contextMenu.target)); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 text-red-600"
              >
                Delete Row {contextMenu.target}
              </button>
              <button 
                type="button"
                onClick={() => { onHideRow(Number(contextMenu.target)); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 border-t border-slate-100"
              >
                Hide Row {contextMenu.target}
              </button>
              <button 
                type="button"
                onClick={() => { onUnhideAllRows(); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2"
              >
                Unhide All Rows
              </button>
            </>
          )}
          {contextMenu.type === 'cell' && (
            <>
              <button 
                type="button"
                onClick={() => { onCopy(); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2"
              >
                Copy (Ctrl+C)
              </button>
              <button 
                type="button"
                onClick={() => { onCut(); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2"
              >
                Cut (Ctrl+X)
              </button>
              <button 
                type="button"
                onClick={() => { onPaste(); setContextMenu(null); }}
                disabled={!hasClipboard}
                className={`w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer flex items-center gap-2 ${!hasClipboard ? 'opacity-40 cursor-not-allowed hover:bg-transparent' : ''}`}
              >
                Paste (Ctrl+V)
              </button>
              <button 
                type="button"
                onClick={() => { onCellChanged(String(contextMenu.target), ""); setContextMenu(null); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 cursor-pointer text-red-600 border-t border-slate-100 font-bold"
              >
                Clear Cell
              </button>
            </>
          )}
        </div>
      )}
    </div>
  );
};
