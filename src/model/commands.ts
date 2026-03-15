import { SpreadsheetModel } from './SpreadsheetModel';
import { CellData, CellRange, CellValue, cellKey } from './types';
import { PasteCell } from '../messages';

export interface EditCommand {
  execute(model: SpreadsheetModel): void;
  undo(model: SpreadsheetModel): void;
  description: string;
}

// ── SetCellValue ────────────────────────────────────────────────────────────

export class SetCellValueCommand implements EditCommand {
  description: string;
  private oldValue: CellValue = null;
  private oldFormula: string | null = null;
  private captured = false;

  constructor(
    private sheetName: string,
    private col: number,
    private row: number,
    private newValue: CellValue,
    private newFormula: string | null,
  ) {
    this.description = `Set cell ${sheetName}!${col},${row}`;
  }

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    // Capture old value only on first execute so undo always restores
    // the state from before this command was ever applied.
    if (!this.captured) {
      const cell = sheet.getCell(this.col, this.row);
      this.oldValue = cell.rawValue;
      this.oldFormula = cell.formula;
      this.captured = true;
    }
    sheet.setCell(this.col, this.row, {
      rawValue: this.newValue,
      formula: this.newFormula,
      computedValue: this.newFormula ? null : this.newValue,
    });
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    sheet.setCell(this.col, this.row, {
      rawValue: this.oldValue,
      formula: this.oldFormula,
      computedValue: this.oldFormula ? null : this.oldValue,
    });
  }
}

// ── SetCellStyle ────────────────────────────────────────────────────────────

export class SetCellStyleCommand implements EditCommand {
  description = 'Set cell style';
  private oldStyleIds: Map<string, string | null> = new Map();

  constructor(
    private sheetName: string,
    private range: CellRange,
    private styleId: string,
  ) {}

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    for (let r = this.range.startRow; r <= this.range.endRow; r++) {
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        const cell = sheet.getCell(c, r);
        this.oldStyleIds.set(cellKey(c, r), cell.styleId);
        sheet.setCell(c, r, { styleId: this.styleId });
      }
    }
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    for (let r = this.range.startRow; r <= this.range.endRow; r++) {
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        const key = cellKey(c, r);
        sheet.setCell(c, r, { styleId: this.oldStyleIds.get(key) ?? null });
      }
    }
  }
}

// ── MergeCells ──────────────────────────────────────────────────────────────

export class MergeCellsCommand implements EditCommand {
  description = 'Merge cells';
  private previousData: Map<string, CellData> = new Map();

  constructor(
    private sheetName: string,
    private range: CellRange,
  ) {}

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    const colSpan = this.range.endCol - this.range.startCol + 1;
    const rowSpan = this.range.endRow - this.range.startRow + 1;

    // Save previous state
    for (let r = this.range.startRow; r <= this.range.endRow; r++) {
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        this.previousData.set(cellKey(c, r), { ...sheet.getCell(c, r) });
      }
    }

    // Set the anchor cell
    sheet.setCell(this.range.startCol, this.range.startRow, {
      mergeColSpan: colSpan,
      mergeRowSpan: rowSpan,
    });

    // Mark other cells as merged
    for (let r = this.range.startRow; r <= this.range.endRow; r++) {
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        if (c === this.range.startCol && r === this.range.startRow) continue;
        sheet.setCell(c, r, {
          mergedInto: { sheet: this.sheetName, col: this.range.startCol, row: this.range.startRow },
          rawValue: null,
          formula: null,
          computedValue: null,
        });
      }
    }
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    for (const [key, data] of this.previousData) {
      const idx = key.indexOf(',');
      const col = parseInt(key.substring(0, idx), 10);
      const row = parseInt(key.substring(idx + 1), 10);
      sheet.setCell(col, row, data);
    }
  }
}

// ── InsertRows/Columns ──────────────────────────────────────────────────────

export class InsertRowsCommand implements EditCommand {
  description: string;
  constructor(
    private sheetName: string,
    private at: number,
    private count: number,
  ) {
    this.description = `Insert ${count} row(s) at ${at}`;
  }
  execute(model: SpreadsheetModel): void {
    model.getSheetByName(this.sheetName)?.insertRows(this.at, this.count);
  }
  undo(model: SpreadsheetModel): void {
    model.getSheetByName(this.sheetName)?.deleteRows(this.at, this.count);
  }
}

export class InsertColumnsCommand implements EditCommand {
  description: string;
  constructor(
    private sheetName: string,
    private at: number,
    private count: number,
  ) {
    this.description = `Insert ${count} column(s) at ${at}`;
  }
  execute(model: SpreadsheetModel): void {
    model.getSheetByName(this.sheetName)?.insertColumns(this.at, this.count);
  }
  undo(model: SpreadsheetModel): void {
    model.getSheetByName(this.sheetName)?.deleteColumns(this.at, this.count);
  }
}

export class DeleteRowsCommand implements EditCommand {
  description: string;
  private savedCells: Array<{ col: number; row: number; data: CellData }> = [];
  private savedHeights: number[] = [];

  constructor(
    private sheetName: string,
    private at: number,
    private count: number,
  ) {
    this.description = `Delete ${count} row(s) at ${at}`;
  }

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    // Save cells in deleted range
    this.savedCells = sheet
      .getAllCells()
      .filter((c) => c.row >= this.at && c.row < this.at + this.count);
    this.savedHeights = Array.from({ length: this.count }, (_, i) =>
      sheet.getRowHeight(this.at + i),
    );
    sheet.deleteRows(this.at, this.count);
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    sheet.insertRows(this.at, this.count);
    // Restore row heights
    for (let i = 0; i < this.savedHeights.length; i++) {
      sheet.setRowHeight(this.at + i, this.savedHeights[i]);
    }
    // Restore cells
    for (const cell of this.savedCells) {
      sheet.setCell(cell.col, cell.row, cell.data);
    }
  }
}

export class DeleteColumnsCommand implements EditCommand {
  description: string;
  private savedCells: Array<{ col: number; row: number; data: CellData }> = [];
  private savedWidths: number[] = [];

  constructor(
    private sheetName: string,
    private at: number,
    private count: number,
  ) {
    this.description = `Delete ${count} column(s) at ${at}`;
  }

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    this.savedCells = sheet
      .getAllCells()
      .filter((c) => c.col >= this.at && c.col < this.at + this.count);
    this.savedWidths = Array.from({ length: this.count }, (_, i) =>
      sheet.getColumnWidth(this.at + i),
    );
    sheet.deleteColumns(this.at, this.count);
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    sheet.insertColumns(this.at, this.count);
    for (let i = 0; i < this.savedWidths.length; i++) {
      sheet.setColumnWidth(this.at + i, this.savedWidths[i]);
    }
    for (const cell of this.savedCells) {
      sheet.setCell(cell.col, cell.row, cell.data);
    }
  }
}

// ── Resize ──────────────────────────────────────────────────────────────────

export class ResizeColumnCommand implements EditCommand {
  description = 'Resize column';
  private oldWidth = 0;

  constructor(
    private sheetName: string,
    private col: number,
    private newWidth: number,
  ) {}

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    this.oldWidth = sheet.getColumnWidth(this.col);
    sheet.setColumnWidth(this.col, this.newWidth);
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    sheet.setColumnWidth(this.col, this.oldWidth);
  }
}

export class ResizeRowCommand implements EditCommand {
  description = 'Resize row';
  private oldHeight = 0;

  constructor(
    private sheetName: string,
    private row: number,
    private newHeight: number,
  ) {}

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    this.oldHeight = sheet.getRowHeight(this.row);
    sheet.setRowHeight(this.row, this.newHeight);
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    sheet.setRowHeight(this.row, this.oldHeight);
  }
}

// ── Sort ────────────────────────────────────────────────────────────────────

export class SortRangeMultiCommand implements EditCommand {
  description = 'Sort range';
  private savedCells: Array<{ col: number; row: number; data: CellData }> = [];

  constructor(
    private sheetName: string,
    private range: CellRange,
    private keys: Array<{ column: number; ascending: boolean }>,
  ) {}

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;

    // Save current state
    this.savedCells = [];
    for (let r = this.range.startRow; r <= this.range.endRow; r++) {
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        if (sheet.hasCell(c, r)) {
          this.savedCells.push({ col: c, row: r, data: { ...sheet.getCell(c, r) } });
        }
      }
    }

    // Collect rows as arrays
    const rows: Array<{ rowIdx: number; cells: Map<number, CellData> }> = [];
    for (let r = this.range.startRow; r <= this.range.endRow; r++) {
      const cellMap = new Map<number, CellData>();
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        if (sheet.hasCell(c, r)) {
          cellMap.set(c, { ...sheet.getCell(c, r) });
        }
      }
      rows.push({ rowIdx: r, cells: cellMap });
    }

    // Sort by multiple keys (stable: sort keys in reverse order for stable multi-key)
    rows.sort((a, b) => {
      for (const key of this.keys) {
        const aCell = a.cells.get(key.column);
        const bCell = b.cells.get(key.column);
        const aVal = aCell?.computedValue ?? aCell?.rawValue ?? null;
        const bVal = bCell?.computedValue ?? bCell?.rawValue ?? null;

        if (aVal === null && bVal === null) continue;
        if (aVal === null) return 1;
        if (bVal === null) return -1;

        let cmp: number;
        if (typeof aVal === 'number' && typeof bVal === 'number') {
          cmp = aVal - bVal;
        } else {
          cmp = String(aVal).localeCompare(String(bVal));
        }
        if (cmp !== 0) return key.ascending ? cmp : -cmp;
      }
      return 0;
    });

    // Write sorted rows back
    for (let i = 0; i < rows.length; i++) {
      const targetRow = this.range.startRow + i;
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        const cellData = rows[i].cells.get(c);
        if (cellData) {
          sheet.setCell(c, targetRow, cellData);
        } else {
          sheet.deleteCell(c, targetRow);
        }
      }
    }
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    // Clear the range
    for (let r = this.range.startRow; r <= this.range.endRow; r++) {
      for (let c = this.range.startCol; c <= this.range.endCol; c++) {
        sheet.deleteCell(c, r);
      }
    }
    // Restore saved cells
    for (const cell of this.savedCells) {
      sheet.setCell(cell.col, cell.row, cell.data);
    }
  }
}

/** Single-key sort (kept for backward compat — delegates to SortRangeMultiCommand) */
export class SortRangeCommand extends SortRangeMultiCommand {
  constructor(sheetName: string, range: CellRange, sortColumn: number, ascending: boolean) {
    super(sheetName, range, [{ column: sortColumn, ascending }]);
  }
}

// ── PasteSpecial ─────────────────────────────────────────────────────────────

export class PasteSpecialCommand implements EditCommand {
  description: string;
  private savedCells: Array<{ col: number; row: number; data: CellData }> = [];

  constructor(
    private sheetName: string,
    private startCol: number,
    private startRow: number,
    private cells: PasteCell[],
    private mode: 'valuesOnly' | 'formatsOnly' | 'all',
  ) {
    this.description = `Paste special (${mode})`;
  }

  execute(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;

    this.savedCells = [];
    for (const item of this.cells) {
      const col = this.startCol + item.col;
      const row = this.startRow + item.row;
      this.savedCells.push({ col, row, data: { ...sheet.getCell(col, row) } });

      if (this.mode === 'valuesOnly') {
        const isFormula = item.value.startsWith('=');
        sheet.setCell(col, row, {
          rawValue: isFormula ? null : item.value === '' ? null : item.value,
          formula: isFormula ? item.value.substring(1) : null,
          computedValue: isFormula ? null : item.value === '' ? null : item.value,
        });
      } else if (this.mode === 'formatsOnly') {
        sheet.setCell(col, row, { styleId: item.styleId });
      } else {
        const isFormula = item.value.startsWith('=');
        sheet.setCell(col, row, {
          rawValue: isFormula ? null : item.value === '' ? null : item.value,
          formula: isFormula ? item.value.substring(1) : null,
          computedValue: isFormula ? null : item.value === '' ? null : item.value,
          styleId: item.styleId,
        });
      }
    }
  }

  undo(model: SpreadsheetModel): void {
    const sheet = model.getSheetByName(this.sheetName);
    if (!sheet) return;
    for (const saved of this.savedCells) {
      sheet.setCell(saved.col, saved.row, saved.data);
    }
  }
}
