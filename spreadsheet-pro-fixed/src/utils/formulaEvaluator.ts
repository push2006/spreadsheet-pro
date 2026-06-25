import { SheetData } from '../types';

// Helper: Convert column letter (e.g. "AA") to index (0-based)
export function colNameToIndex(col: string): number {
  let ans = 0;
  const cleaned = col.toUpperCase();
  for (let i = 0; i < cleaned.length; i++) {
    ans = ans * 26 + (cleaned.charCodeAt(i) - 64);
  }
  return ans - 1;
}

// Helper: Convert index (0-based) to column letter (e.g. "AA")
export function indexToColName(index: number): string {
  let name = "";
  let n = index;
  do {
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return name;
}

// Helper: Parse cell coordinate (e.g., "A1" into { col: "A", colIndex: 0, rowIndex: 1 })
export function parseCellId(cellId: string) {
  const match = cellId.toUpperCase().match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  return {
    col: match[1],
    colIndex: colNameToIndex(match[1]),
    rowIndex: parseInt(match[2], 10),
  };
}

// Helper: Evaluate a cell value, with circular reference guard
export function getCellEvaluatedValue(
  cellId: string,
  sheetData: SheetData,
  evaluatedCache: { [cellId: string]: string },
  runningStack: Set<string> = new Set()
): string {
  const normalizedId = cellId.toUpperCase();
  
  if (runningStack.has(normalizedId)) {
    return "#CIRCULAR!";
  }
  
  if (evaluatedCache[normalizedId] !== undefined) {
    return evaluatedCache[normalizedId];
  }
  
  const rawValue = sheetData[normalizedId];
  if (rawValue === undefined || rawValue === "") {
    return "";
  }
  
  if (typeof rawValue === "string" && (rawValue.startsWith("=") || rawValue.startsWith("+"))) {
    runningStack.add(normalizedId);
    const result = evaluateFormula(rawValue, sheetData, evaluatedCache, runningStack);
    runningStack.delete(normalizedId);
    evaluatedCache[normalizedId] = result;
    return result;
  }
  
  return rawValue;
}

// Helper: Retrieve evaluating array of values inside range e.g. "A1:B3"
export function getRangeValuesRaw(
  rangeStr: string,
  sheetData: SheetData,
  evaluatedCache: { [cellId: string]: string },
  runningStack: Set<string>
): string[] {
  const parts = rangeStr.split(":");
  if (parts.length !== 2) return [];
  
  const startParsed = parseCellId(parts[0]);
  const endParsed = parseCellId(parts[1]);
  if (!startParsed || !endParsed) return [];
  
  const minCol = Math.min(startParsed.colIndex, endParsed.colIndex);
  const maxCol = Math.max(startParsed.colIndex, endParsed.colIndex);
  const minRow = Math.min(startParsed.rowIndex, endParsed.rowIndex);
  const maxRow = Math.max(startParsed.rowIndex, endParsed.rowIndex);
  
  const values: string[] = [];
  for (let r = minRow; r <= maxRow; r++) {
    for (let c = minCol; c <= maxCol; c++) {
      const cellId = indexToColName(c) + r;
      values.push(getCellEvaluatedValue(cellId, sheetData, evaluatedCache, runningStack));
    }
  }
  return values;
}

// Helper: Sum range values
export function getRangeValuesDecimal(
  rangeStr: string,
  sheetData: SheetData,
  evaluatedCache: { [cellId: string]: string },
  runningStack: Set<string>
): number[] {
  const raw = getRangeValuesRaw(rangeStr, sheetData, evaluatedCache, runningStack);
  return raw.map(v => parseFloat(v)).filter(v => !isNaN(v));
}

// Main evaluation logic
export function evaluateFormula(
  formula: string,
  sheetData: SheetData,
  evaluatedCache: { [cellId: string]: string } = {},
  runningStack: Set<string> = new Set()
): string {
  try {
    let cleanFormula = formula.trim();
    if (!cleanFormula.startsWith("=") && !cleanFormula.startsWith("+")) {
      return cleanFormula;
    }
    
    // Remove the leading '=' or '+'
    let expr = cleanFormula.substring(1);
    
    // Evaluate functions from the innermost to the outermost.
    // An innermost function call stands for FUNCTION_NAME( arguments_without_parentheses )
    // We execute replacements inside a loop until no more function calls can be simplified.
    let loopCount = 0;
    const MAX_LOOPS = 50; // Prevention against infinite loops
    
    const functionRegex = /(SUM|AVERAGE|MIN|MAX|COUNT|ABS|SQRT|ROUND|LEN|UPPER|LOWER|LEFT|RIGHT|CONCAT|IF|AND|OR|NOT|TODAY|NOW|YEAR|MONTH|DAY|DAYS|DATE|DATEDIF|MATCH|INDEX|VLOOKUP|XLOOKUP|SUMIF|COUNTIF|AVERAGEIF|ISBLANK|ISNUMBER)\(([^()]+)?\)/i;
    
    while (functionRegex.test(expr) && loopCount < MAX_LOOPS) {
      expr = expr.replace(functionRegex, (match, funcName: string, argString: string) => {
        const uFuncName = funcName.toUpperCase();
        const argsStr = argString || "";
        const args = argsStr.split(",").map(a => a.trim());
        
        switch (uFuncName) {
          case "SUM": {
            if (args[0]?.includes(":")) {
              const vals = getRangeValuesDecimal(args[0], sheetData, evaluatedCache, runningStack);
              return String(vals.reduce((a, b) => a + b, 0));
            }
            const sumVals = args.map(a => parseFloat(a)).filter(n => !isNaN(n));
            return String(sumVals.reduce((a, b) => a + b, 0));
          }
          
          case "AVERAGE": {
            if (args[0]?.includes(":")) {
              const vals = getRangeValuesDecimal(args[0], sheetData, evaluatedCache, runningStack);
              if (vals.length === 0) return "0";
              return String(vals.reduce((a, b) => a + b, 0) / vals.length);
            }
            const avgVals = args.map(a => parseFloat(a)).filter(n => !isNaN(n));
            if (avgVals.length === 0) return "0";
            return String(avgVals.reduce((a, b) => a + b, 0) / avgVals.length);
          }
          
          case "MIN": {
            if (args[0]?.includes(":")) {
              const vals = getRangeValuesDecimal(args[0], sheetData, evaluatedCache, runningStack);
              if (vals.length === 0) return "0";
              return String(Math.min(...vals));
            }
            const minVals = args.map(a => parseFloat(a)).filter(n => !isNaN(n));
            if (minVals.length === 0) return "0";
            return String(Math.min(...minVals));
          }
          
          case "MAX": {
            if (args[0]?.includes(":")) {
              const vals = getRangeValuesDecimal(args[0], sheetData, evaluatedCache, runningStack);
              if (vals.length === 0) return "0";
              return String(Math.max(...vals));
            }
            const maxVals = args.map(a => parseFloat(a)).filter(n => !isNaN(n));
            if (maxVals.length === 0) return "0";
            return String(Math.max(...maxVals));
          }
          
          case "COUNT": {
            if (args[0]?.includes(":")) {
              const vals = getRangeValuesDecimal(args[0], sheetData, evaluatedCache, runningStack);
              return String(vals.length);
            }
            const countVals = args.map(a => parseFloat(a)).filter(n => !isNaN(n));
            return String(countVals.length);
          }
          
          case "ABS": {
            const val = parseFloat(args[0]);
            return isNaN(val) ? "0" : String(Math.abs(val));
          }
          
          case "SQRT": {
            const val = parseFloat(args[0]);
            return isNaN(val) ? "0" : String(Math.sqrt(val));
          }
          
          case "ROUND": {
            const val = parseFloat(args[0]);
            const decimals = args[1] ? parseInt(args[1], 10) : 0;
            if (isNaN(val)) return "0";
            return String(val.toFixed(decimals));
          }
          
          case "LEN": {
            // strip surrounding quotes if literal text
            const raw = args[0] || "";
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            return String(cleanStr.length);
          }
          
          case "UPPER": {
            const raw = args.join(",");
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            return `"${cleanStr.toUpperCase()}"`;
          }
          
          case "LOWER": {
            const raw = args.join(",");
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            return `"${cleanStr.toLowerCase()}"`;
          }
          
          case "LEFT": {
            const raw = args[0] || "";
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            const size = args[1] ? parseInt(args[1], 10) : 1;
            return `"${cleanStr.substring(0, size)}"`;
          }
          
          case "RIGHT": {
            const raw = args[0] || "";
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            const size = args[1] ? parseInt(args[1], 10) : 1;
            return `"${cleanStr.substring(cleanStr.length - size)}"`;
          }
          
          case "CONCAT": {
            const cleanArgs = args.map(a => (a.startsWith('"') && a.endsWith('"') ? a.slice(1, -1) : a));
            return `"${cleanArgs.join("")}"`;
          }
          
          case "IF": {
            const conditionStr = args[0] || "false";
            const trueVal = args[1] || "";
            const falseVal = args[2] || "";
            
            // Evaluate condition as safely as possible in simple environment
            let conditionResult = false;
            try {
              // replace simple strings or operators safely
              const parsedCond = conditionStr
                .replace(/"([^"]*)"/g, "'$1'") // swap double with single quotes
                .trim();
              const evalFn = new Function(`return (${parsedCond})`);
              conditionResult = !!evalFn();
            } catch {
              conditionResult = false;
            }
            
            const resultVal = conditionResult ? trueVal : falseVal;
            return resultVal;
          }
          
          case "AND": {
            const evaluations = args.map(v => {
              try {
                const evalFn = new Function(`return (${v})`);
                return !!evalFn();
              } catch {
                return false;
              }
            });
            return String(evaluations.every(b => b));
          }
          
          case "OR": {
            const evaluations = args.map(v => {
              try {
                const evalFn = new Function(`return (${v})`);
                return !!evalFn();
              } catch {
                return false;
              }
            });
            return String(evaluations.some(b => b));
          }
          
          case "NOT": {
            try {
              const evalFn = new Function(`return (${args[0]})`);
              return String(!evalFn());
            } catch {
              return "false";
            }
          }
          
          case "TODAY": {
            return `"${new Date().toISOString().split("T")[0]}"`;
          }
          
          case "NOW": {
            return `"${new Date().toLocaleString()}"`;
          }
          
          case "YEAR": {
            const raw = args[0] || "";
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            return String(new Date(cleanStr).getFullYear());
          }
          
          case "MONTH": {
            const raw = args[0] || "";
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            return String(new Date(cleanStr).getMonth() + 1);
          }
          
          case "DAY": {
            const raw = args[0] || "";
            const cleanStr = raw.startsWith('"') && raw.endsWith('"') ? raw.slice(1, -1) : raw;
            return String(new Date(cleanStr).getDate());
          }

          case "DAYS": {
            const end = (args[0] || "").replace(/"/g, "").trim();
            const start = (args[1] || "").replace(/"/g, "").trim();
            const diffTime = new Date(end).getTime() - new Date(start).getTime();
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            return isNaN(diffDays) ? "0" : String(diffDays);
          }
          
          case "DATE": {
            const y = parseInt(args[0] || "0", 10);
            const m = parseInt(args[1] || "1", 10);
            const d = parseInt(args[2] || "1", 10);
            const date = new Date(y, m - 1, d);
            const yyyy = date.getFullYear();
            const mm = String(date.getMonth() + 1).padStart(2, '0');
            const dd = String(date.getDate()).padStart(2, '0');
            return `"${yyyy}-${mm}-${dd}"`;
          }

          case "DATEDIF": {
            const start = new Date((args[0] || "").replace(/"/g, "").trim());
            const end = new Date((args[1] || "").replace(/"/g, "").trim());
            const unit = (args[2] || "D").replace(/"/g, "").toUpperCase().trim();
            if (isNaN(start.getTime()) || isNaN(end.getTime())) return "0";
            
            if (unit === "Y") {
              return String(end.getFullYear() - start.getFullYear());
            } else if (unit === "M") {
              const months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
              return String(months);
            } else {
              const diffTime = end.getTime() - start.getTime();
              return String(Math.floor(diffTime / (1000 * 60 * 60 * 24)));
            }
          }
          
          case "MATCH": {
            const lookupValue = (args[0] || "").replace(/"/g, "").trim();
            const rangeStr = args[1] || "";
            const values = getRangeValuesRaw(rangeStr, sheetData, evaluatedCache, runningStack);
            
            const idx = values.findIndex(v => String(v).replace(/"/g, "").trim() === lookupValue);
            return idx >= 0 ? String(idx + 1) : "#N/A";
          }
          
          case "INDEX": {
            const rangeStr = args[0] || "";
            const pos = parseInt(args[1] || "1", 10);
            const values = getRangeValuesRaw(rangeStr, sheetData, evaluatedCache, runningStack);
            return values[pos - 1] !== undefined ? values[pos - 1] : "#N/A";
          }
          
          case "VLOOKUP": {
            const lookupValue = (args[0] || "").replace(/"/g, "").trim();
            const tableRange = args[1] || "";
            const colIndex = parseInt(args[2] || "1", 10);
            
            // Split tableRange e.g. "A1:C10"
            const tblParts = tableRange.split(":");
            if (tblParts.length !== 2) return "#N/A";
            
            const startNode = parseCellId(tblParts[0]);
            const endNode = parseCellId(tblParts[1]);
            if (!startNode || !endNode) return "#N/A";
            
            // Loop through each row in table
            const startCol = Math.min(startNode.colIndex, endNode.colIndex);
            const startRow = Math.min(startNode.rowIndex, endNode.rowIndex);
            const endRow = Math.max(startNode.rowIndex, endNode.rowIndex);
            
            for (let r = startRow; r <= endRow; r++) {
              // Leftmost column cell in this row of the table
              const firstCellId = indexToColName(startCol) + r;
              const firstVal = getCellEvaluatedValue(firstCellId, sheetData, evaluatedCache, runningStack).replace(/"/g, "").trim();
              
              if (String(firstVal) === lookupValue) {
                const targetColIndex = startCol + colIndex - 1;
                const targetCellId = indexToColName(targetColIndex) + r;
                return getCellEvaluatedValue(targetCellId, sheetData, evaluatedCache, runningStack);
              }
            }
            return "#N/A";
          }
          
          case "XLOOKUP": {
            const lookupValue = (args[0] || "").replace(/"/g, "").trim();
            const lookupRange = args[1] || "";
            const returnRange = args[2] || "";
            
            const lookupVals = getRangeValuesRaw(lookupRange, sheetData, evaluatedCache, runningStack);
            const returnVals = getRangeValuesRaw(returnRange, sheetData, evaluatedCache, runningStack);
            
            const idx = lookupVals.findIndex(v => String(v).replace(/"/g, "").trim() === lookupValue);
            return idx >= 0 && returnVals[idx] !== undefined ? returnVals[idx] : "#N/A";
          }
          
          case "SUMIF": {
            const criteriaRange = args[0] || "";
            const criteria = (args[1] || "").replace(/"/g, "").trim();
            const sumRange = args[2] || criteriaRange; // defaults to criteria range if omitted
            
            const checks = getRangeValuesRaw(criteriaRange, sheetData, evaluatedCache, runningStack);
            const sums = getRangeValuesRaw(sumRange, sheetData, evaluatedCache, runningStack);
            
            let total = 0;
            checks.forEach((v, idx) => {
              const cleanV = String(v).replace(/"/g, "").trim();
              const matches = checkCriteriaMatch(cleanV, criteria);
              if (matches) {
                const val = parseFloat(sums[idx]);
                if (!isNaN(val)) total += val;
              }
            });
            return String(total);
          }
          
          case "COUNTIF": {
            const rangeStr = args[0] || "";
            const criteria = (args[1] || "").replace(/"/g, "").trim();
            const vals = getRangeValuesRaw(rangeStr, sheetData, evaluatedCache, runningStack);
            
            const count = vals.filter(v => {
              const cleanV = String(v).replace(/"/g, "").trim();
              return checkCriteriaMatch(cleanV, criteria);
            }).length;
            return String(count);
          }
          
          case "AVERAGEIF": {
            const criteriaRange = args[0] || "";
            const criteria = (args[1] || "").replace(/"/g, "").trim();
            const avgRange = args[2] || criteriaRange;
            
            const checks = getRangeValuesRaw(criteriaRange, sheetData, evaluatedCache, runningStack);
            const avgs = getRangeValuesRaw(avgRange, sheetData, evaluatedCache, runningStack);
            
            let total = 0;
            let count = 0;
            checks.forEach((v, idx) => {
              const cleanV = String(v).replace(/"/g, "").trim();
              if (checkCriteriaMatch(cleanV, criteria)) {
                const val = parseFloat(avgs[idx]);
                if (!isNaN(val)) {
                  total += val;
                  count++;
                }
              }
            });
            return count > 0 ? String(total / count) : "0";
          }
          
          case "ISBLANK": {
            const cellId = (args[0] || "").trim();
            const raw = sheetData[cellId.toUpperCase()];
            return String(raw === undefined || raw === "");
          }
          
          case "ISNUMBER": {
            const val = parseFloat(args[0]);
            return String(!isNaN(val));
          }
          
          default:
            return "0";
        }
      });
      loopCount++;
    }
    
    // Evaluate pure cell references e.g. A1, B2 if they are not in function signatures.
    // Ensure we don't accidentally match part of strings (avoid matching alphabetical texts inside double quotes).
    // Simple way is to find standalone cell IDs in expression.Stand-alone regex represents boundary and cells:
    const cellRefRegex = /\b([A-Z]+)(\d+)\b/g;
    expr = expr.replace(cellRefRegex, (match, col, row) => {
      const cellId = col.toUpperCase() + row;
      const val = getCellEvaluatedValue(cellId, sheetData, evaluatedCache, runningStack);
      return val;
    });
    
    // Finally, evaluate any basic math operators (+, -, *, /, %, parenthesis), stripping any trailing quotes or literal markers.
    let mathExpr = expr
      .replace(/"/g, "") // strip quotes for raw math operations
      .trim();
      
    if (mathExpr === "") {
      return "";
    }
    
    // Safety guard for math calculation
    if (/^[0-9.+\-*/%() ]+$/.test(mathExpr)) {
      try {
        const result = new Function(`return (${mathExpr})`)();
        return result !== undefined ? String(result) : "";
      } catch {
        return "#VALUE!";
      }
    }
    
    return expr; // return raw text string if it's alphanumeric string e.g. "HELLO"
  } catch (err) {
    return "#ERROR";
  }
}

// Helpers for conditional matches: supports operators like ">10", "<=5", "<>0", or exact matches
function checkCriteriaMatch(val: string, criteria: string): boolean {
  if (criteria.startsWith(">=")) {
    const lim = parseFloat(criteria.substring(2));
    const v = parseFloat(val);
    return !isNaN(v) && !isNaN(lim) ? v >= lim : val >= criteria.substring(2);
  }
  if (criteria.startsWith("<=")) {
    const lim = parseFloat(criteria.substring(2));
    const v = parseFloat(val);
    return !isNaN(v) && !isNaN(lim) ? v <= lim : val <= criteria.substring(2);
  }
  if (criteria.startsWith(">")) {
    const lim = parseFloat(criteria.substring(1));
    const v = parseFloat(val);
    return !isNaN(v) && !isNaN(lim) ? v > lim : val > criteria.substring(1);
  }
  if (criteria.startsWith("<")) {
    const lim = parseFloat(criteria.substring(1));
    const v = parseFloat(val);
    return !isNaN(v) && !isNaN(lim) ? v < lim : val < criteria.substring(1);
  }
  if (criteria.startsWith("<>")) {
    return val !== criteria.substring(2);
  }
  return val === criteria;
}
