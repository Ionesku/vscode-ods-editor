import { FormulaError } from '../model/types';

// ── Tokens ──────────────────────────────────────────────────────────────────

export enum TokenType {
  Number = 'Number',
  String = 'String',
  Boolean = 'Boolean',
  CellRef = 'CellRef',
  FunctionName = 'FunctionName',
  Operator = 'Operator',
  LParen = 'LParen',
  RParen = 'RParen',
  Comma = 'Comma',
  Colon = 'Colon',
  Semicolon = 'Semicolon',
  EOF = 'EOF',
}

export interface Token {
  type: TokenType;
  value: string;
  position: number;
}

// ── AST Nodes ───────────────────────────────────────────────────────────────

export type ASTNode =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'error'; errorType: FormulaError }
  | { type: 'cellRef'; sheet?: string; col: number; row: number; absCol: boolean; absRow: boolean }
  | {
      type: 'rangeRef';
      sheet?: string;
      startCol: number;
      startRow: number;
      endCol: number;
      endRow: number;
    }
  | { type: 'namedRangeRef'; name: string }
  | { type: 'binaryOp'; op: string; left: ASTNode; right: ASTNode }
  | { type: 'unaryOp'; op: string; operand: ASTNode }
  | { type: 'functionCall'; name: string; args: ASTNode[] };

// ── Cell values for formula evaluation ──────────────────────────────────────

export type CellValueType = number | string | boolean | null | FormulaError;

/** A 2D array representing a range of cell values */
export type RangeValue = CellValueType[][];

/** Value that a formula function can receive */
export type FormulaValue = CellValueType | RangeValue;

/** Check if a value is a range (2D array) */
export function isRangeValue(val: FormulaValue): val is RangeValue {
  return Array.isArray(val) && (val.length === 0 || Array.isArray(val[0]));
}

/** Check if a value is a formula error */
export function isFormulaError(val: unknown): val is FormulaError {
  return (
    (typeof val === 'string' && val.startsWith('#') && val.endsWith('!')) ||
    val === '#NAME?' ||
    val === '#N/A'
  );
}

/** Flatten a range to a flat array of values */
export function flattenRange(val: FormulaValue): CellValueType[] {
  if (isRangeValue(val)) {
    const result: CellValueType[] = [];
    for (const row of val) {
      for (const cell of row) {
        result.push(cell);
      }
    }
    return result;
  }
  return [val];
}

/** Convert a value to number, or return error */
export function toNumber(val: CellValueType): number | FormulaError {
  if (typeof val === 'number') return val;
  if (typeof val === 'boolean') return val ? 1 : 0;
  if (val === null || val === '') return 0;
  if (typeof val === 'string') {
    const n = Number(val);
    if (!isNaN(n)) return n;
    return FormulaError.VALUE;
  }
  return val as FormulaError;
}
