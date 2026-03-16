import {
  CellData,
  CellRange,
  DataValidation,
  FilterCriteria,
  SortSpec,
  cellKey,
  createEmptyCell,
} from './types';
import { ConditionalFormatRule } from '../messages';

export const DEFAULT_COL_WIDTH = 80;
export const DEFAULT_ROW_HEIGHT = 24;
const DEFAULT_COL_COUNT = 1024;
const DEFAULT_ROW_COUNT = 65536;

export class SheetModel {
  name: string;
  private cells: Map<string, CellData> = new Map();
  /** Sparse map of non-default column widths. Missing entries use DEFAULT_COL_WIDTH. */
  private _columnWidths: Map<number, number> = new Map();
  /** Sparse map of non-default row heights. Missing entries use DEFAULT_ROW_HEIGHT. */
  private _rowHeights: Map<number, number> = new Map();
  columnCount: number;
  rowCount: number;
  filters: Map<number, FilterCriteria> = new Map();
  sortSpec: SortSpec | null = null;
  hiddenRows: Set<number> = new Set();
  frozenRows = 0;
  frozenCols = 0;
  conditionalFormats: ConditionalFormatRule[] = [];
  dataValidations: DataValidation[] = [];

  constructor(name: string, colCount = DEFAULT_COL_COUNT, rowCount = DEFAULT_ROW_COUNT) {
    this.name = name;
    this.columnCount = colCount;
    this.rowCount = rowCount;
  }

  getColumnWidth(col: number): number {
    return this._columnWidths.get(col) ?? DEFAULT_COL_WIDTH;
  }

  getRowHeight(row: number): number {
    return this._rowHeights.get(row) ?? DEFAULT_ROW_HEIGHT;
  }

  /**
   * Returns a sparse array of non-default column widths up to `maxCol` (exclusive).
   * Positions with default width are omitted (undefined); callers should use
   * `?? DEFAULT_COL_WIDTH` when reading individual entries.
   */
  getColumnWidthsArray(maxCol: number): number[] {
    const arr: number[] = new Array(maxCol);
    for (let i = 0; i < maxCol; i++) {
      arr[i] = this._columnWidths.get(i) ?? DEFAULT_COL_WIDTH;
    }
    return arr;
  }

  getRowHeightsArray(maxRow: number): number[] {
    const arr: number[] = new Array(maxRow);
    for (let i = 0; i < maxRow; i++) {
      arr[i] = this._rowHeights.get(i) ?? DEFAULT_ROW_HEIGHT;
    }
    return arr;
  }

  /** Iterate over all explicitly-set (non-default) column widths */
  get nonDefaultColumnWidths(): ReadonlyMap<number, number> {
    return this._columnWidths;
  }

  /** Iterate over all explicitly-set (non-default) row heights */
  get nonDefaultRowHeights(): ReadonlyMap<number, number> {
    return this._rowHeights;
  }

  getCell(col: number, row: number): CellData {
    const key = cellKey(col, row);
    return this.cells.get(key) ?? createEmptyCell();
  }

  setCell(col: number, row: number, data: Partial<CellData>): void {
    const key = cellKey(col, row);
    const existing = this.cells.get(key) ?? createEmptyCell();
    const updated = { ...existing, ...data };

    // If the cell is now effectively empty, remove it to save memory
    if (
      updated.rawValue === null &&
      updated.formula === null &&
      updated.styleId === null &&
      updated.mergeColSpan === 1 &&
      updated.mergeRowSpan === 1 &&
      updated.mergedInto === null &&
      !updated.comment
    ) {
      this.cells.delete(key);
    } else {
      this.cells.set(key, updated);
    }
  }

  hasCell(col: number, row: number): boolean {
    return this.cells.has(cellKey(col, row));
  }

  deleteCell(col: number, row: number): void {
    this.cells.delete(cellKey(col, row));
  }

  getCellRange(range: CellRange): CellData[][] {
    const result: CellData[][] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row: CellData[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(this.getCell(c, r));
      }
      result.push(row);
    }
    return result;
  }

  /** Get all non-empty cell entries */
  getAllCells(): Array<{ col: number; row: number; data: CellData }> {
    const result: Array<{ col: number; row: number; data: CellData }> = [];
    for (const [key, data] of this.cells) {
      const idx = key.indexOf(',');
      const col = parseInt(key.substring(0, idx), 10);
      const row = parseInt(key.substring(idx + 1), 10);
      result.push({ col, row, data });
    }
    return result;
  }

  /** Get the used range (bounding box of all non-empty cells) */
  getUsedRange(): CellRange | null {
    let minCol = Infinity,
      maxCol = -1,
      minRow = Infinity,
      maxRow = -1;
    for (const [key] of this.cells) {
      const idx = key.indexOf(',');
      const col = parseInt(key.substring(0, idx), 10);
      const row = parseInt(key.substring(idx + 1), 10);
      if (col < minCol) minCol = col;
      if (col > maxCol) maxCol = col;
      if (row < minRow) minRow = row;
      if (row > maxRow) maxRow = row;
    }
    if (maxCol === -1) return null;
    return {
      sheet: this.name,
      startCol: minCol,
      startRow: minRow,
      endCol: maxCol,
      endRow: maxRow,
    };
  }

  get cellCount(): number {
    return this.cells.size;
  }

  setColumnWidth(col: number, width: number): void {
    if (col >= 0 && col < this.columnCount) {
      const w = Math.max(20, width);
      if (w === DEFAULT_COL_WIDTH) {
        this._columnWidths.delete(col);
      } else {
        this._columnWidths.set(col, w);
      }
    }
  }

  setRowHeight(row: number, height: number): void {
    if (row >= 0 && row < this.rowCount) {
      const h = Math.max(10, height);
      if (h === DEFAULT_ROW_HEIGHT) {
        this._rowHeights.delete(row);
      } else {
        this._rowHeights.set(row, h);
      }
    }
  }

  /** Insert rows at the given position */
  insertRows(at: number, count: number): void {
    // Shift existing cell data
    const toMove: Array<{ col: number; row: number; data: CellData }> = [];
    for (const entry of this.getAllCells()) {
      if (entry.row >= at) {
        toMove.push(entry);
      }
    }
    for (const entry of toMove) {
      this.cells.delete(cellKey(entry.col, entry.row));
    }
    for (const entry of toMove) {
      this.cells.set(cellKey(entry.col, entry.row + count), entry.data);
    }
    // Shift row heights upward
    const newHeights: Map<number, number> = new Map();
    for (const [r, h] of this._rowHeights) {
      if (r < at) {
        newHeights.set(r, h);
      } else {
        newHeights.set(r + count, h);
      }
    }
    this._rowHeights = newHeights;
    this.rowCount += count;
  }

  /** Insert columns at the given position */
  insertColumns(at: number, count: number): void {
    const toMove: Array<{ col: number; row: number; data: CellData }> = [];
    for (const entry of this.getAllCells()) {
      if (entry.col >= at) {
        toMove.push(entry);
      }
    }
    for (const entry of toMove) {
      this.cells.delete(cellKey(entry.col, entry.row));
    }
    for (const entry of toMove) {
      this.cells.set(cellKey(entry.col + count, entry.row), entry.data);
    }
    // Shift column widths
    const newWidths: Map<number, number> = new Map();
    for (const [c, w] of this._columnWidths) {
      if (c < at) {
        newWidths.set(c, w);
      } else {
        newWidths.set(c + count, w);
      }
    }
    this._columnWidths = newWidths;
    this.columnCount += count;
  }

  /** Delete rows */
  deleteRows(at: number, count: number): void {
    for (const entry of this.getAllCells()) {
      if (entry.row >= at && entry.row < at + count) {
        this.cells.delete(cellKey(entry.col, entry.row));
      }
    }
    const toMove: Array<{ col: number; row: number; data: CellData }> = [];
    for (const entry of this.getAllCells()) {
      if (entry.row >= at + count) {
        toMove.push(entry);
      }
    }
    for (const entry of toMove) {
      this.cells.delete(cellKey(entry.col, entry.row));
    }
    for (const entry of toMove) {
      this.cells.set(cellKey(entry.col, entry.row - count), entry.data);
    }
    // Shift row heights
    const newHeights: Map<number, number> = new Map();
    for (const [r, h] of this._rowHeights) {
      if (r < at) {
        newHeights.set(r, h);
      } else if (r >= at + count) {
        newHeights.set(r - count, h);
      }
      // rows in [at, at+count) are deleted
    }
    this._rowHeights = newHeights;
    this.rowCount -= count;
  }

  /** Delete columns */
  deleteColumns(at: number, count: number): void {
    for (const entry of this.getAllCells()) {
      if (entry.col >= at && entry.col < at + count) {
        this.cells.delete(cellKey(entry.col, entry.row));
      }
    }
    const toMove: Array<{ col: number; row: number; data: CellData }> = [];
    for (const entry of this.getAllCells()) {
      if (entry.col >= at + count) {
        toMove.push(entry);
      }
    }
    for (const entry of toMove) {
      this.cells.delete(cellKey(entry.col, entry.row));
    }
    for (const entry of toMove) {
      this.cells.set(cellKey(entry.col - count, entry.row), entry.data);
    }
    // Shift column widths
    const newWidths: Map<number, number> = new Map();
    for (const [c, w] of this._columnWidths) {
      if (c < at) {
        newWidths.set(c, w);
      } else if (c >= at + count) {
        newWidths.set(c - count, w);
      }
    }
    this._columnWidths = newWidths;
    this.columnCount -= count;
  }
}
