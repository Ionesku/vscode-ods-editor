// ── Cell addressing ──────────────────────────────────────────────────────────

export interface CellAddress {
  sheet: string;
  col: number;
  row: number;
}

export interface CellRange {
  sheet: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
}

// ── Cell values ─────────────────────────────────────────────────────────────

export enum FormulaError {
  REF = '#REF!',
  VALUE = '#VALUE!',
  DIV0 = '#DIV/0!',
  NAME = '#NAME?',
  NA = '#N/A',
  NUM = '#NUM!',
  NULL = '#NULL!',
  CIRC = '#CIRC!',
}

export type CellValue = number | string | boolean | null | FormulaError;

export enum CellType {
  Empty = 'empty',
  Number = 'number',
  String = 'string',
  Boolean = 'boolean',
  Formula = 'formula',
  Error = 'error',
}

// ── Cell data ───────────────────────────────────────────────────────────────

export interface CellData {
  rawValue: CellValue;
  formula: string | null;
  computedValue: CellValue;
  styleId: string | null;
  mergeColSpan: number;
  mergeRowSpan: number;
  mergedInto: CellAddress | null;
}

export function createEmptyCell(): CellData {
  return {
    rawValue: null,
    formula: null,
    computedValue: null,
    styleId: null,
    mergeColSpan: 1,
    mergeRowSpan: 1,
    mergedInto: null,
  };
}

// ── Styles ──────────────────────────────────────────────────────────────────

export interface BorderStyle {
  width: 'thin' | 'medium' | 'thick';
  style: 'solid' | 'dashed' | 'dotted';
  color: string;
}

export interface CellStyle {
  id: string;
  fontFamily?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  textColor?: string;
  backgroundColor?: string;
  horizontalAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  borderTop?: BorderStyle;
  borderRight?: BorderStyle;
  borderBottom?: BorderStyle;
  borderLeft?: BorderStyle;
  wrapText?: boolean;
  numberFormat?: string;
}

// ── Sorting / Filtering ─────────────────────────────────────────────────────

export interface SortSpec {
  column: number;
  ascending: boolean;
}

export interface FilterCriteria {
  column: number;
  values: Set<string>; // allowed values (empty = show all)
  customFilter?: string; // regex or expression
}

// ── Named Ranges ─────────────────────────────────────────────────────────────

export interface NamedRange {
  name: string;
  sheet: string;
  range: CellRange;
}

// ── Data Validation ──────────────────────────────────────────────────────────

export interface DataValidation {
  id: string;
  range: CellRange;
  /** 'list' shows a dropdown; 'none' removes validation */
  type: 'list' | 'none';
  /** For type='list': comma-separated values or a range address like "Sheet1.A1:A5" */
  listSource: string;
  /** Optional message shown when the user enters an invalid value */
  errorMessage?: string;
}

// ── Document metadata ───────────────────────────────────────────────────────

export interface DocumentMetadata {
  creator?: string;
  creationDate?: string;
  modifiedDate?: string;
  title?: string;
  description?: string;
}

// ── Helpers ─────────────────────────────────────────────────────────────────

/** Convert 0-based column index to letter(s): 0->A, 25->Z, 26->AA */
export function colToLetter(col: number): string {
  let result = '';
  let c = col;
  while (c >= 0) {
    result = String.fromCharCode((c % 26) + 65) + result;
    c = Math.floor(c / 26) - 1;
  }
  return result;
}

/** Convert column letter(s) to 0-based index: A->0, Z->25, AA->26 */
export function letterToCol(letters: string): number {
  let result = 0;
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64);
  }
  return result - 1;
}

/** Create a cell key string for Map storage */
export function cellKey(col: number, row: number): string {
  return `${col},${row}`;
}

/** Parse a cell key string back to col, row */
export function parseCellKey(key: string): { col: number; row: number } {
  const idx = key.indexOf(',');
  return {
    col: parseInt(key.substring(0, idx), 10),
    row: parseInt(key.substring(idx + 1), 10),
  };
}
