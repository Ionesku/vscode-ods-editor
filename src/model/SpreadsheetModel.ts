import { SheetModel, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './SheetModel';
import { CellStyle, DataValidation, DocumentMetadata, NamedRange } from './types';

export class SpreadsheetModel {
  sheets: SheetModel[] = [];
  styles: Map<string, CellStyle> = new Map();
  metadata: DocumentMetadata = {};
  activeSheetIndex: number = 0;
  namedRanges: Map<string, NamedRange> = new Map();

  /** Raw XML data for elements we don't parse (for round-trip fidelity) */
  rawExtras: Map<string, string> = new Map();

  constructor() {
    // Create a default sheet
    this.sheets.push(new SheetModel('Sheet1'));
  }

  get activeSheet(): SheetModel {
    return this.sheets[this.activeSheetIndex];
  }

  addSheet(name: string): SheetModel {
    const sheet = new SheetModel(this.uniqueSheetName(name));
    this.sheets.push(sheet);
    return sheet;
  }

  /** Returns a sheet name that doesn't conflict with existing sheets. */
  private uniqueSheetName(base: string): string {
    const existing = new Set(this.sheets.map((s) => s.name));
    if (!existing.has(base)) return base;
    let i = 2;
    while (existing.has(`${base} (${i})`)) i++;
    return `${base} (${i})`;
  }

  removeSheet(index: number): void {
    if (this.sheets.length <= 1) return; // Cannot remove last sheet
    this.sheets.splice(index, 1);
    if (this.activeSheetIndex >= this.sheets.length) {
      this.activeSheetIndex = this.sheets.length - 1;
    }
  }

  renameSheet(index: number, name: string): boolean {
    if (index < 0 || index >= this.sheets.length) return false;
    const trimmed = name.trim();
    if (!trimmed) return false;
    // Reject if another sheet already has this name
    if (this.sheets.some((s, i) => i !== index && s.name === trimmed)) return false;
    this.sheets[index].name = trimmed;
    return true;
  }

  reorderSheet(fromIndex: number, toIndex: number): void {
    if (fromIndex < 0 || fromIndex >= this.sheets.length) return;
    if (toIndex < 0 || toIndex >= this.sheets.length) return;
    const [sheet] = this.sheets.splice(fromIndex, 1);
    this.sheets.splice(toIndex, 0, sheet);
    // Adjust active sheet index
    if (this.activeSheetIndex === fromIndex) {
      this.activeSheetIndex = toIndex;
    } else {
      if (fromIndex < this.activeSheetIndex && toIndex >= this.activeSheetIndex) {
        this.activeSheetIndex--;
      } else if (fromIndex > this.activeSheetIndex && toIndex <= this.activeSheetIndex) {
        this.activeSheetIndex++;
      }
    }
  }

  getSheetByName(name: string): SheetModel | undefined {
    return this.sheets.find((s) => s.name === name);
  }

  getOrCreateStyle(style: Partial<CellStyle> & { id: string }): CellStyle {
    const existing = this.styles.get(style.id);
    if (existing) return existing;
    const { id, ...rest } = style;
    const full: CellStyle = { id, ...rest };
    this.styles.set(id, full);
    return full;
  }

  resolveStyle(styleId: string | null): CellStyle | undefined {
    if (!styleId) return undefined;
    return this.styles.get(styleId);
  }

  defineNamedRange(name: string, sheet: string, range: import('./types').CellRange): void {
    this.namedRanges.set(name.toUpperCase(), { name, sheet, range });
  }

  deleteNamedRange(name: string): void {
    this.namedRanges.delete(name.toUpperCase());
  }

  resolveNamedRange(name: string): NamedRange | undefined {
    return this.namedRanges.get(name.toUpperCase());
  }

  /** Serialize model to a plain object for webview transfer */
  serialize(): SerializedModel {
    return {
      sheets: this.sheets.map((sheet) => {
        const used = sheet.getUsedRange();
        const minCols = Math.max(used ? used.endCol + 2 : 0, sheet.frozenCols + 1, 26);
        const minRows = Math.max(used ? used.endRow + 2 : 0, sheet.frozenRows + 1, 50);

        // Find the last column/row with a non-default size to avoid shipping large arrays
        let lastCustomCol = minCols - 1;
        for (const [c] of sheet.nonDefaultColumnWidths) {
          if (c > lastCustomCol) lastCustomCol = c;
        }
        let lastCustomRow = minRows - 1;
        for (const [r] of sheet.nonDefaultRowHeights) {
          if (r > lastCustomRow) lastCustomRow = r;
        }

        return {
          name: sheet.name,
          columnCount: sheet.columnCount,
          rowCount: sheet.rowCount,
          // Only send up to the last non-default value; webview pads the rest with defaults
          columnWidths: sheet.getColumnWidthsArray(lastCustomCol + 1),
          rowHeights: sheet.getRowHeightsArray(lastCustomRow + 1),
          cells: sheet.getAllCells().map((c) => ({
            col: c.col,
            row: c.row,
            data: c.data,
          })),
          frozenRows: sheet.frozenRows,
          frozenCols: sheet.frozenCols,
          conditionalFormats: sheet.conditionalFormats,
          activeFilters: Array.from(sheet.filters.entries()).map(([col, f]) => ({
            column: col,
            values: Array.from(f.values),
          })),
          hiddenRows: Array.from(sheet.hiddenRows),
          dataValidations: sheet.dataValidations,
        };
      }),
      styles: Object.fromEntries(this.styles),
      activeSheetIndex: this.activeSheetIndex,
      metadata: this.metadata,
      namedRanges: Array.from(this.namedRanges.values()),
    };
  }
}

export interface SerializedSheet {
  name: string;
  columnCount: number;
  rowCount: number;
  columnWidths: number[];
  rowHeights: number[];
  cells: Array<{ col: number; row: number; data: import('./types').CellData }>;
  frozenRows: number;
  frozenCols: number;
  conditionalFormats?: import('../messages').ConditionalFormatRule[];
  activeFilters?: Array<{ column: number; values: string[] }>;
  hiddenRows?: number[];
  dataValidations?: DataValidation[];
}

export interface SerializedModel {
  sheets: SerializedSheet[];
  styles: Record<string, CellStyle>;
  activeSheetIndex: number;
  metadata: DocumentMetadata;
  namedRanges?: NamedRange[];
}
