import {
  ASTNode,
  CellValueType,
  FormulaValue,
  RangeValue,
  isFormulaError,
  isRangeValue,
} from './types';
import { FormulaError, NamedRange } from '../model/types';
import { FunctionRegistry } from './functions/index';

export type CellValueGetter = (sheet: string, col: number, row: number) => CellValueType;
export type NamedRangeResolver = (name: string) => NamedRange | undefined;

/** The address of the cell currently being evaluated (for ROW/COLUMN/OFFSET context) */
export interface EvalContext {
  sheet: string;
  col: number;
  row: number;
}

export class Evaluator {
  /** Set before each cell evaluation so ROW()/COLUMN()/OFFSET() can read the current address */
  currentCell: EvalContext = { sheet: '', col: 0, row: 0 };

  /** Read a cell value by address — used by OFFSET, INDIRECT */
  getCellAt(sheet: string, col: number, row: number): CellValueType {
    return this.getCellValue(sheet, col, row);
  }

  constructor(
    private getCellValue: CellValueGetter,
    private functionRegistry: FunctionRegistry,
    private resolveNamedRange: NamedRangeResolver = () => undefined,
  ) {}

  evaluate(node: ASTNode, contextSheet: string): FormulaValue {
    switch (node.type) {
      case 'number':
        return node.value;
      case 'string':
        return node.value;
      case 'boolean':
        return node.value;
      case 'error':
        return node.errorType;

      case 'cellRef': {
        const sheet = node.sheet ?? contextSheet;
        const val = this.getCellValue(sheet, node.col, node.row);
        return val;
      }

      case 'namedRangeRef': {
        const nr = this.resolveNamedRange(node.name);
        if (!nr) return FormulaError.NAME;
        const r = nr.range;
        if (r.startCol === r.endCol && r.startRow === r.endRow) {
          // Single cell named range — return the cell value
          return this.getCellValue(nr.sheet, r.startCol, r.startRow);
        }
        // Multi-cell named range — return as range value
        const rows: CellValueType[][] = [];
        for (let row = r.startRow; row <= r.endRow; row++) {
          const rowVals: CellValueType[] = [];
          for (let col = r.startCol; col <= r.endCol; col++) {
            rowVals.push(this.getCellValue(nr.sheet, col, row));
          }
          rows.push(rowVals);
        }
        return rows;
      }

      case 'rangeRef': {
        const sheet = node.sheet ?? contextSheet;
        const rows: CellValueType[][] = [];
        for (let r = node.startRow; r <= node.endRow; r++) {
          const row: CellValueType[] = [];
          for (let c = node.startCol; c <= node.endCol; c++) {
            row.push(this.getCellValue(sheet, c, r));
          }
          rows.push(row);
        }
        return rows;
      }

      case 'binaryOp':
        return this.evaluateBinaryOp(node.op, node.left, node.right, contextSheet);

      case 'unaryOp':
        return this.evaluateUnaryOp(node.op, node.operand, contextSheet);

      case 'functionCall':
        return this.evaluateFunction(node.name, node.args, contextSheet);

      default:
        return FormulaError.VALUE;
    }
  }

  private evaluateBinaryOp(op: string, left: ASTNode, right: ASTNode, sheet: string): FormulaValue {
    const lval = this.evaluate(left, sheet);
    const rval = this.evaluate(right, sheet);

    // Check for errors
    if (isFormulaError(lval)) return lval;
    if (isFormulaError(rval)) return rval;

    // Element-wise array operations (CSE / dynamic array behaviour)
    if (isRangeValue(lval) || isRangeValue(rval)) {
      return this.evaluateBinaryOpElementWise(op, lval as FormulaValue, rval as FormulaValue);
    }

    // Concatenation
    if (op === '&') {
      return String(lval ?? '') + String(rval ?? '');
    }

    // Numeric operations
    const lnum = this.toNumber(lval);
    const rnum = this.toNumber(rval);

    if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
      return this.compareValues(op, lval, rval);
    }

    if (isFormulaError(lnum)) return lnum;
    if (isFormulaError(rnum)) return rnum;

    switch (op) {
      case '+':
        return lnum + rnum;
      case '-':
        return lnum - rnum;
      case '*':
        return lnum * rnum;
      case '/':
        if (rnum === 0) return FormulaError.DIV0;
        return lnum / rnum;
      case '^':
        return Math.pow(lnum, rnum);
      default:
        return FormulaError.VALUE;
    }
  }

  /** Element-wise binary operation on two values where at least one is a RangeValue */
  private evaluateBinaryOpElementWise(
    op: string,
    lval: FormulaValue,
    rval: FormulaValue,
  ): RangeValue {
    // Normalise scalars to 1×1 arrays for uniform iteration
    const la: CellValueType[][] = isRangeValue(lval)
      ? (lval as RangeValue)
      : [[lval as CellValueType]];
    const ra: CellValueType[][] = isRangeValue(rval)
      ? (rval as RangeValue)
      : [[rval as CellValueType]];
    const rows = Math.max(la.length, ra.length);
    const cols = Math.max(la[0]?.length ?? 1, ra[0]?.length ?? 1);
    const result: CellValueType[][] = [];
    for (let r = 0; r < rows; r++) {
      const row: CellValueType[] = [];
      for (let c = 0; c < cols; c++) {
        // Out-of-bounds on a 1-row/1-col operand means broadcast that row/col
        const lv = la[Math.min(r, la.length - 1)]?.[Math.min(c, (la[0]?.length ?? 1) - 1)] ?? null;
        const rv = ra[Math.min(r, ra.length - 1)]?.[Math.min(c, (ra[0]?.length ?? 1) - 1)] ?? null;
        if (isFormulaError(lv)) {
          row.push(lv);
          continue;
        }
        if (isFormulaError(rv)) {
          row.push(rv);
          continue;
        }
        if (op === '&') {
          row.push(String(lv ?? '') + String(rv ?? ''));
          continue;
        }
        if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
          row.push(this.compareValues(op, lv, rv));
          continue;
        }
        const ln = this.toNumber(lv as CellValueType);
        const rn = this.toNumber(rv as CellValueType);
        if (isFormulaError(ln)) {
          row.push(ln);
          continue;
        }
        if (isFormulaError(rn)) {
          row.push(rn);
          continue;
        }
        switch (op) {
          case '+':
            row.push(ln + rn);
            break;
          case '-':
            row.push(ln - rn);
            break;
          case '*':
            row.push(ln * rn);
            break;
          case '/':
            row.push(rn === 0 ? FormulaError.DIV0 : ln / rn);
            break;
          case '^':
            row.push(Math.pow(ln, rn));
            break;
          default:
            row.push(FormulaError.VALUE);
        }
      }
      result.push(row);
    }
    return result;
  }

  private compareValues(op: string, lval: FormulaValue, rval: FormulaValue): boolean {
    // For comparison, use raw values
    const l = lval === null ? '' : lval;
    const r = rval === null ? '' : rval;

    if (typeof l === 'number' && typeof r === 'number') {
      switch (op) {
        case '=':
          return l === r;
        case '<>':
          return l !== r;
        case '<':
          return l < r;
        case '>':
          return l > r;
        case '<=':
          return l <= r;
        case '>=':
          return l >= r;
      }
    }

    const ls = String(l).toLowerCase();
    const rs = String(r).toLowerCase();
    switch (op) {
      case '=':
        return ls === rs;
      case '<>':
        return ls !== rs;
      case '<':
        return ls < rs;
      case '>':
        return ls > rs;
      case '<=':
        return ls <= rs;
      case '>=':
        return ls >= rs;
    }
    return false;
  }

  private evaluateUnaryOp(op: string, operand: ASTNode, sheet: string): CellValueType {
    const val = this.evaluate(operand, sheet);
    if (isFormulaError(val)) return val as FormulaError;
    const num = this.toNumber(val);
    if (isFormulaError(num)) return num;
    if (op === '-') return -num;
    return num;
  }

  private evaluateFunction(name: string, args: ASTNode[], sheet: string): FormulaValue {
    const fn = this.functionRegistry.get(name);
    if (!fn) return FormulaError.NAME;

    // Special handling for IF — lazy evaluation
    if (name === 'IF') {
      if (args.length < 2) return FormulaError.VALUE;
      const condition = this.evaluate(args[0], sheet);
      if (isFormulaError(condition)) return condition;
      const truthiness = this.isTruthy(condition);
      if (truthiness) {
        return this.evaluate(args[1], sheet);
      } else {
        return args.length >= 3 ? this.evaluate(args[2], sheet) : false;
      }
    }

    // Special handling for IFERROR
    if (name === 'IFERROR') {
      if (args.length < 2) return FormulaError.VALUE;
      const val = this.evaluate(args[0], sheet);
      if (isFormulaError(val)) {
        return this.evaluate(args[1], sheet);
      }
      return val;
    }

    // Evaluate all arguments
    const evaluatedArgs = args.map((arg) => this.evaluate(arg, sheet));
    return fn(evaluatedArgs, this);
  }

  private isTruthy(val: FormulaValue): boolean {
    if (typeof val === 'boolean') return val;
    if (typeof val === 'number') return val !== 0;
    if (typeof val === 'string') return val.length > 0;
    return val !== null;
  }

  toNumber(val: FormulaValue): number | FormulaError {
    if (typeof val === 'number') return val;
    if (typeof val === 'boolean') return val ? 1 : 0;
    if (val === null || val === '') return 0;
    if (typeof val === 'string') {
      const n = Number(val);
      if (!isNaN(n)) return n;
      return FormulaError.VALUE;
    }
    if (isFormulaError(val)) return val;
    return FormulaError.VALUE;
  }
}
