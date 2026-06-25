export interface CellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // Text color CSS, hex or tailwind class
  bg?: string;    // Background color hex
  align?: 'left' | 'center' | 'right';
  fontSize?: number; // In px, e.g., 12, 14, 16
  format?: 'general' | 'number' | 'currency' | 'percent' | 'scientific';
  border?: 'none' | 'all' | 'bottom' | 'top' | 'thick' | 'double-bottom';
}

export interface SheetData {
  [cellId: string]: string; // raw typed values/formulas, e.g., "A1": "=SUM(B1:B10)" or "A1": "42"
}

export interface SheetStyles {
  [cellId: string]: CellStyle;
}

export interface MergedRange {
  start: string; // e.g. "A1"
  end: string;   // e.g. "C3"
  cells: string[]; // cell IDs included, e.g., ["A1", "A2", "A3", "B1", "B2", "B3", "C1", "C2", "C3"]
}

export interface Sheet {
  id: string;
  name: string;
  data: SheetData;
  styles: SheetStyles;
  columnWidths: { [col: string]: number };
  rowHeights: { [row: number]: number };
  mergedCells: MergedRange[];
  hiddenColumns?: { [col: string]: boolean };
  hiddenRows?: { [row: number]: boolean };
}

export interface WorkbookState {
  sheets: { [id: string]: Sheet };
  sheetIds: string[];
  activeSheetId: string;
}

export interface SelectionRange {
  start: string; // cellId, e.g., "A1"
  end: string;   // cellId, e.g., "C3"
}

export interface ClipboardData {
  value: string;
  style: CellStyle | null;
}

export interface FloatingWidget {
  id: string;
  type: 'chart' | 'note' | 'shape' | 'slicer' | 'pivot' | 'text' | 'wordart';
  x: number;
  y: number;
  sheetId: string;
  title: string;
  content?: string;
  range?: string; // e.g. A1:B5
  chartType?: 'bar' | 'line' | 'area';
  shapeType?: 'rect' | 'circle' | 'arrow';
  color?: string;
  pivotRows?: string;
  pivotCols?: string;
  pivotVals?: string;
  pivotOp?: 'sum' | 'avg' | 'count' | 'max' | 'min';
}
