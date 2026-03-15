import { CellData, CellStyle, CellRange, createEmptyCell, cellKey } from '../../shared/types';
import { ConditionalFormatRule } from '../../shared/messages';

export interface SheetState {
  name: string;
  columnCount: number;
  rowCount: number;
  columnWidths: number[];
  rowHeights: number[];
  cells: Map<string, CellData>;
  hiddenRows: Set<number>;
  frozenRows: number;
  frozenCols: number;
  conditionalFormats: ConditionalFormatRule[];
  activeFilters: Map<number, Set<string>>; // col -> allowed values
}

export class WebviewState {
  sheets: SheetState[] = [];
  styles: Map<string, CellStyle> = new Map();
  activeSheetIndex = 0;

  // Selection
  selectedCol = 0;
  selectedRow = 0;
  selectionEndCol = 0;
  selectionEndRow = 0;
  isSelecting = false;

  // Scroll
  scrollX = 0;
  scrollY = 0;

  // Editing
  isEditing = false;
  editValue = '';

  get activeSheet(): SheetState | undefined {
    return this.sheets[this.activeSheetIndex];
  }

  get selectionRange(): CellRange | null {
    const sheet = this.activeSheet;
    if (!sheet) return null;
    return {
      sheet: sheet.name,
      startCol: Math.min(this.selectedCol, this.selectionEndCol),
      startRow: Math.min(this.selectedRow, this.selectionEndRow),
      endCol: Math.max(this.selectedCol, this.selectionEndCol),
      endRow: Math.max(this.selectedRow, this.selectionEndRow),
    };
  }

  loadFromSerialized(data: {
    sheets: Array<{
      name: string;
      columnCount: number;
      rowCount: number;
      columnWidths: number[];
      rowHeights: number[];
      cells: Array<{ col: number; row: number; data: CellData }>;
      frozenRows?: number;
      frozenCols?: number;
      conditionalFormats?: ConditionalFormatRule[];
      activeFilters?: Array<{ column: number; values: string[] }>;
      hiddenRows?: number[];
    }>;
    styles: Record<string, CellStyle>;
    activeSheetIndex: number;
  }): void {
    const DEFAULT_COL_WIDTH = 80;
    const DEFAULT_ROW_HEIGHT = 24;
    this.sheets = data.sheets.map((s) => {
      const cellMap = new Map<string, CellData>();
      for (const c of s.cells) {
        cellMap.set(cellKey(c.col, c.row), c.data);
      }

      // Expand sparse width/height arrays back to full size using defaults
      const columnWidths = s.columnWidths.slice();
      while (columnWidths.length < s.columnCount) columnWidths.push(DEFAULT_COL_WIDTH);
      const rowHeights = s.rowHeights.slice();
      while (rowHeights.length < s.rowCount) rowHeights.push(DEFAULT_ROW_HEIGHT);

      return {
        name: s.name,
        columnCount: s.columnCount,
        rowCount: s.rowCount,
        columnWidths,
        rowHeights,
        cells: cellMap,
        hiddenRows: new Set<number>(s.hiddenRows ?? []),
        frozenRows: s.frozenRows ?? 0,
        frozenCols: s.frozenCols ?? 0,
        conditionalFormats: s.conditionalFormats ?? [],
        activeFilters: new Map((s.activeFilters ?? []).map((f) => [f.column, new Set(f.values)])),
      };
    });
    this.styles = new Map(Object.entries(data.styles));
    this.activeSheetIndex = data.activeSheetIndex;
  }

  getCell(col: number, row: number): CellData {
    const sheet = this.activeSheet;
    if (!sheet) return createEmptyCell();
    return sheet.cells.get(cellKey(col, row)) ?? createEmptyCell();
  }

  getCellDisplay(col: number, row: number): string {
    const cell = this.getCell(col, row);
    const value = cell.computedValue ?? cell.rawValue;
    if (value === null) return '';
    if (typeof value === 'number') {
      const style = this.getCellStyle(col, row);
      if (style?.numberFormat) return formatNumber(value, style.numberFormat);
      return String(value);
    }
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    return String(value);
  }

  /** Format a number using the given format string — used for live preview */
  formatForPreview(value: number, fmt: string): string {
    return formatNumber(value, fmt);
  }

  getCellStyle(col: number, row: number): CellStyle | undefined {
    const cell = this.getCell(col, row);
    const baseStyle = cell.styleId ? this.styles.get(cell.styleId) : undefined;

    // Apply conditional formatting
    const sheet = this.activeSheet;
    if (!sheet || sheet.conditionalFormats.length === 0) return baseStyle;

    for (const rule of sheet.conditionalFormats) {
      if (col < rule.range.startCol || col > rule.range.endCol) continue;
      if (row < rule.range.startRow || row > rule.range.endRow) continue;

      const val = cell.computedValue ?? cell.rawValue;
      if (evaluateCondition(rule, val)) {
        // Merge conditional style on top of base style
        return { ...(baseStyle ?? { id: '' }), ...rule.style, id: baseStyle?.id ?? '' };
      }
    }

    return baseStyle;
  }
}

function formatNumber(value: number, fmt: string): string {
  if (!fmt || fmt === 'General' || fmt === 'general') return String(value);

  // Date format — contains yyyy / mm / dd / hh
  if (/yyyy|yy|mm|dd|hh/i.test(fmt) && !fmt.includes('#') && !fmt.includes('0')) {
    return formatDateSerial(value, fmt);
  }

  // Percent format
  if (fmt.includes('%')) {
    const pctDecimals = fmt.includes('.') ? (fmt.split('.')[1]?.replace('%', '').length ?? 0) : 0;
    return (value * 100).toFixed(pctDecimals) + '%';
  }

  // Scientific notation  0.00E+00
  if (/E[+-]/i.test(fmt)) {
    const decMatch = fmt.match(/\.([0#]+)E/i);
    const sciDecimals = decMatch ? decMatch[1].length : 2;
    return value.toExponential(sciDecimals).toUpperCase();
  }

  // Currency prefix ($, €, £, ¥, ₽)
  const currencyMatch = fmt.match(/^([$€£¥₽])/);
  const currency = currencyMatch ? currencyMatch[1] : '';
  const cleanFmt = currency ? fmt.substring(currency.length) : fmt;

  // Count decimal places
  const decPart = cleanFmt.split('.')[1];
  const decimals = decPart ? decPart.replace(/[^0#]/g, '').length : 0;

  // Thousands separator
  const useThousands = cleanFmt.includes(',');

  let result: string;
  if (useThousands) {
    const fixed = Math.abs(value).toFixed(decimals);
    const [intPart, fracPart] = fixed.split('.');
    const withCommas = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    result = fracPart ? withCommas + '.' + fracPart : withCommas;
    if (value < 0) result = '-' + result;
  } else {
    result = value.toFixed(decimals);
  }

  return currency + result;
}

/** Convert an Excel serial date (days since 1900-01-01, 1-based, with 1900 leap year bug) to a JS Date */
function serialToDate(serial: number): Date {
  // Excel day 1 = 1900-01-01; serial 60 is the fake 1900-02-29 (Excel bug), skip it
  const adjusted = serial > 59 ? serial - 1 : serial;
  return new Date(Date.UTC(1899, 11, 30) + adjusted * 86400000);
}

function formatDateSerial(value: number, fmt: string): string {
  const d = serialToDate(value);
  if (isNaN(d.getTime())) return String(value);

  const yyyy = String(d.getUTCFullYear()).padStart(4, '0');
  const yy = yyyy.slice(2);
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const hh = String(d.getUTCHours()).padStart(2, '0');
  const min = String(d.getUTCMinutes()).padStart(2, '0');
  const ss = String(d.getUTCSeconds()).padStart(2, '0');

  return fmt
    .replace(/yyyy/gi, yyyy)
    .replace(/yy/gi, yy)
    .replace(/mm/gi, mm)
    .replace(/dd/gi, dd)
    .replace(/hh/gi, hh)
    .replace(/mi|nn/gi, min) // mi or nn for minutes (to avoid conflict with mm=month)
    .replace(/ss/gi, ss);
}

function evaluateCondition(rule: ConditionalFormatRule, val: unknown): boolean {
  const strVal = val !== null && val !== undefined ? String(val) : '';
  const numVal = typeof val === 'number' ? val : Number(strVal);
  const v1 = Number(rule.value1);

  switch (rule.condition) {
    case 'greaterThan':
      return !isNaN(numVal) && !isNaN(v1) && numVal > v1;
    case 'lessThan':
      return !isNaN(numVal) && !isNaN(v1) && numVal < v1;
    case 'equals':
      return strVal === rule.value1 || (!isNaN(numVal) && numVal === v1);
    case 'notEquals':
      return strVal !== rule.value1 && (isNaN(numVal) || numVal !== v1);
    case 'contains':
      return strVal.toLowerCase().includes(rule.value1.toLowerCase());
    case 'notContains':
      return !strVal.toLowerCase().includes(rule.value1.toLowerCase());
    case 'between': {
      const v2 = Number(rule.value2);
      return !isNaN(numVal) && !isNaN(v1) && !isNaN(v2) && numVal >= v1 && numVal <= v2;
    }
    case 'empty':
      return val === null || val === undefined || strVal === '';
    case 'notEmpty':
      return val !== null && val !== undefined && strVal !== '';
    default:
      return false;
  }
}
