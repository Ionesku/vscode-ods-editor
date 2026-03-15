import { FunctionRegistry } from './index';
import {
  FormulaValue,
  isRangeValue,
  flattenRange,
  isFormulaError,
  toNumber,
  CellValueType,
  RangeValue,
} from '../types';
import { FormulaError } from '../../model/types';
import { letterToCol } from '../../model/types';

export function registerLookupFunctions(registry: FunctionRegistry): void {
  registry.register('VLOOKUP', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const lookupVal = flattenRange(args[0])[0];
    if (isFormulaError(lookupVal)) return lookupVal;
    if (!isRangeValue(args[1])) return FormulaError.VALUE;
    const table = args[1] as RangeValue;
    const colIdx = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(colIdx)) return colIdx;
    const col = Math.floor(colIdx) - 1; // 1-based to 0-based
    const exactMatch = args.length >= 4 ? !flattenRange(args[3])[0] : false; // default: approximate

    if (col < 0 || table.length === 0 || col >= (table[0]?.length ?? 0)) {
      return FormulaError.REF;
    }

    if (exactMatch) {
      for (const row of table) {
        if (valuesEqual(row[0], lookupVal)) {
          return row[col] ?? null;
        }
      }
      return FormulaError.NA;
    } else {
      // Approximate match: find largest value <= lookupVal
      let bestIdx = -1;
      for (let i = 0; i < table.length; i++) {
        const v = table[i][0];
        if (v !== null && compareValues(v, lookupVal) <= 0) {
          bestIdx = i;
        }
      }
      if (bestIdx === -1) return FormulaError.NA;
      return table[bestIdx][col] ?? null;
    }
  });

  registry.register('HLOOKUP', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const lookupVal = flattenRange(args[0])[0];
    if (isFormulaError(lookupVal)) return lookupVal;
    if (!isRangeValue(args[1])) return FormulaError.VALUE;
    const table = args[1] as RangeValue;
    const rowIdx = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(rowIdx)) return rowIdx;
    const row = Math.floor(rowIdx) - 1;
    const exactMatch = args.length >= 4 ? !flattenRange(args[3])[0] : false;

    if (row < 0 || row >= table.length || table[0]?.length === 0) {
      return FormulaError.REF;
    }

    const headerRow = table[0];
    if (exactMatch) {
      for (let c = 0; c < headerRow.length; c++) {
        if (valuesEqual(headerRow[c], lookupVal)) {
          return table[row]?.[c] ?? null;
        }
      }
      return FormulaError.NA;
    } else {
      let bestCol = -1;
      for (let c = 0; c < headerRow.length; c++) {
        if (headerRow[c] !== null && compareValues(headerRow[c], lookupVal) <= 0) {
          bestCol = c;
        }
      }
      if (bestCol === -1) return FormulaError.NA;
      return table[row]?.[bestCol] ?? null;
    }
  });

  registry.register('INDEX', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    if (!isRangeValue(args[0])) return FormulaError.VALUE;
    const range = args[0] as RangeValue;
    const rowNum = toNumber(flattenRange(args[1])[0]);
    if (isFormulaError(rowNum)) return rowNum;
    const row = Math.floor(rowNum) - 1;

    let col = 0;
    if (args.length >= 3) {
      const colNum = toNumber(flattenRange(args[2])[0]);
      if (isFormulaError(colNum)) return colNum;
      col = Math.floor(colNum) - 1;
    }

    if (row < 0 || row >= range.length) return FormulaError.REF;
    if (col < 0 || col >= (range[0]?.length ?? 0)) return FormulaError.REF;
    return range[row]?.[col] ?? null;
  });

  registry.register('MATCH', (args) => {
    if (args.length < 2) return FormulaError.VALUE;
    const lookupVal = flattenRange(args[0])[0];
    if (isFormulaError(lookupVal)) return lookupVal;
    const lookupArray = flattenRange(args[1]);
    const matchType = args.length >= 3 ? toNumber(flattenRange(args[2])[0]) : 1;
    if (isFormulaError(matchType)) return matchType;

    if (matchType === 0) {
      // Exact match
      for (let i = 0; i < lookupArray.length; i++) {
        if (valuesEqual(lookupArray[i], lookupVal)) {
          return i + 1; // 1-based
        }
      }
      return FormulaError.NA;
    } else if (matchType === 1) {
      // Find largest value <= lookupVal (array must be ascending)
      let bestIdx = -1;
      for (let i = 0; i < lookupArray.length; i++) {
        if (lookupArray[i] !== null && compareValues(lookupArray[i], lookupVal) <= 0) {
          bestIdx = i;
        }
      }
      return bestIdx >= 0 ? bestIdx + 1 : FormulaError.NA;
    } else {
      // Find smallest value >= lookupVal (array must be descending)
      let bestIdx = -1;
      for (let i = 0; i < lookupArray.length; i++) {
        if (lookupArray[i] !== null && compareValues(lookupArray[i], lookupVal) >= 0) {
          bestIdx = i;
        }
      }
      return bestIdx >= 0 ? bestIdx + 1 : FormulaError.NA;
    }
  });

  registry.register('ROW', (args, evaluator) => {
    if (args.length === 0) {
      // ROW() — row of the current cell (1-based)
      return evaluator.currentCell.row + 1;
    }
    if (isRangeValue(args[0])) {
      // ROW(range) — row number of the first row of the range
      // We can't easily know the start row from a RangeValue; fall back to current cell
      return evaluator.currentCell.row + 1;
    }
    return evaluator.currentCell.row + 1;
  });

  registry.register('COLUMN', (args, evaluator) => {
    if (args.length === 0) {
      return evaluator.currentCell.col + 1;
    }
    if (isRangeValue(args[0])) {
      return evaluator.currentCell.col + 1;
    }
    return evaluator.currentCell.col + 1;
  });

  registry.register('ROWS', (args) => {
    if (args.length < 1 || !isRangeValue(args[0])) return FormulaError.VALUE;
    return (args[0] as RangeValue).length;
  });

  registry.register('COLUMNS', (args) => {
    if (args.length < 1 || !isRangeValue(args[0])) return FormulaError.VALUE;
    const range = args[0] as RangeValue;
    return range.length > 0 ? range[0].length : 0;
  });

  /**
   * OFFSET(reference, rows, cols, [height], [width])
   * Returns the value of a cell offset from a base reference.
   * When reference is omitted or a range, uses the current cell as base.
   */
  registry.register('OFFSET', (args, evaluator) => {
    if (args.length < 3) return FormulaError.VALUE;
    const rowOffset = toNumber(flattenRange(args[1])[0]);
    const colOffset = toNumber(flattenRange(args[2])[0]);
    if (isFormulaError(rowOffset)) return rowOffset;
    if (isFormulaError(colOffset)) return colOffset;

    // Base reference: use current cell if arg[0] is a range value (already resolved)
    // For a single-cell base we would normally receive the cell value, not its address.
    // Since the evaluator resolves cell references before passing to functions, we use
    // the current cell as the base anchor.
    const baseCol = evaluator.currentCell.col + Math.round(colOffset);
    const baseRow = evaluator.currentCell.row + Math.round(rowOffset);

    if (baseCol < 0 || baseRow < 0) return FormulaError.REF;

    // Height/width (default 1×1)
    const height = args.length >= 4 ? Math.round(Number(flattenRange(args[3])[0]) || 1) : 1;
    const width = args.length >= 5 ? Math.round(Number(flattenRange(args[4])[0]) || 1) : 1;

    if (height === 1 && width === 1) {
      return evaluator.getCellAt(evaluator.currentCell.sheet, baseCol, baseRow);
    }

    // Multi-cell OFFSET — return as range value
    const rows: CellValueType[][] = [];
    for (let r = 0; r < height; r++) {
      const row: CellValueType[] = [];
      for (let c = 0; c < width; c++) {
        row.push(evaluator.getCellAt(evaluator.currentCell.sheet, baseCol + c, baseRow + r));
      }
      rows.push(row);
    }
    return rows;
  });

  /**
   * INDIRECT(ref_text, [a1])
   * Returns the cell value referenced by a string like "A1" or "Sheet1.B2".
   */
  registry.register('INDIRECT', (args, evaluator) => {
    if (args.length < 1) return FormulaError.VALUE;
    const refText = String(flattenRange(args[0])[0] ?? '').trim();
    if (!refText) return FormulaError.REF;

    // Parse A1-style reference, optionally with sheet prefix (Sheet1.A1 or Sheet1!A1)
    let sheetName = evaluator.currentCell.sheet;
    let cellPart = refText;
    const dotIdx = refText.indexOf('.');
    const bangIdx = refText.indexOf('!');
    const sepIdx = dotIdx >= 0 ? dotIdx : bangIdx;
    if (sepIdx > 0) {
      sheetName = refText.substring(0, sepIdx);
      cellPart = refText.substring(sepIdx + 1);
    }

    const match = cellPart.match(/^\$?([A-Za-z]{1,3})\$?(\d+)$/);
    if (!match) return FormulaError.REF;
    const col = letterToCol(match[1].toUpperCase());
    const row = parseInt(match[2], 10) - 1;
    if (col < 0 || row < 0) return FormulaError.REF;

    // Access via the evaluator's getCellValue getter (exposed through getCellAt helper)
    return evaluator.getCellAt(sheetName, col, row);
  });

  /**
   * XLOOKUP(lookup_value, lookup_array, return_array, [not_found], [match_mode], [search_mode])
   * match_mode: 0=exact, -1=exact or smaller, 1=exact or larger, 2=wildcard
   */
  registry.register('XLOOKUP', (args) => {
    if (args.length < 3) return FormulaError.VALUE;
    const lookupVal = flattenRange(args[0])[0];
    if (isFormulaError(lookupVal)) return lookupVal;

    const lookupArr = flattenRange(args[1]);
    const returnArr = flattenRange(args[2]);
    const notFound: FormulaValue = args.length >= 4 ? args[3] : FormulaError.NA;
    const matchMode = args.length >= 5 ? Math.round(Number(flattenRange(args[4])[0]) || 0) : 0;

    let foundIdx = -1;

    if (matchMode === 0) {
      // Exact match
      foundIdx = lookupArr.findIndex((v) => valuesEqual(v, lookupVal));
    } else if (matchMode === -1) {
      // Exact or next smaller
      let bestVal: CellValueType = null;
      for (let i = 0; i < lookupArr.length; i++) {
        const v = lookupArr[i];
        if (v === null) continue;
        if (compareValues(v, lookupVal) <= 0) {
          if (bestVal === null || compareValues(v, bestVal) > 0) {
            bestVal = v;
            foundIdx = i;
          }
        }
      }
    } else if (matchMode === 1) {
      // Exact or next larger
      let bestVal: CellValueType = null;
      for (let i = 0; i < lookupArr.length; i++) {
        const v = lookupArr[i];
        if (v === null) continue;
        if (compareValues(v, lookupVal) >= 0) {
          if (bestVal === null || compareValues(v, bestVal) < 0) {
            bestVal = v;
            foundIdx = i;
          }
        }
      }
    } else if (matchMode === 2) {
      // Wildcard match
      const pattern = String(lookupVal ?? '')
        .toLowerCase()
        .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
        .replace(/\\\*/g, '.*')
        .replace(/\\\?/g, '.');
      const regex = new RegExp('^' + pattern + '$', 'i');
      foundIdx = lookupArr.findIndex((v) => regex.test(String(v ?? '')));
    }

    if (foundIdx < 0 || foundIdx >= returnArr.length) return notFound as CellValueType;
    return returnArr[foundIdx] ?? null;
  });
}

function valuesEqual(a: CellValueType, b: CellValueType): boolean {
  if (a === b) return true;
  if (typeof a === 'string' && typeof b === 'string') {
    return a.toLowerCase() === b.toLowerCase();
  }
  return false;
}

function compareValues(a: CellValueType, b: CellValueType): number {
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  return String(a ?? '').localeCompare(String(b ?? ''), undefined, { sensitivity: 'base' });
}
