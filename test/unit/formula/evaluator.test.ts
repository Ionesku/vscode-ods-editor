import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../src/formula/Tokenizer';
import { Parser } from '../../../src/formula/Parser';
import { Evaluator, CellValueGetter } from '../../../src/formula/Evaluator';
import { createDefaultRegistry } from '../../../src/formula/functions/index';
import { FormulaError } from '../../../src/model/types';
import { CellValueType } from '../../../src/formula/types';

describe('Evaluator', () => {
  const tokenizer = new Tokenizer();
  const parser = new Parser();
  const registry = createDefaultRegistry();

  function evaluate(formula: string, cells: Record<string, CellValueType> = {}): CellValueType {
    const getCellValue: CellValueGetter = (sheet, col, row) => {
      const key = `${String.fromCharCode(65 + col)}${row + 1}`;
      return cells[key] ?? null;
    };
    const evaluator = new Evaluator(getCellValue, registry);
    const tokens = tokenizer.tokenize(formula);
    const ast = parser.parse(tokens);
    const result = evaluator.evaluate(ast, 'Sheet1');
    // Unwrap range values
    if (Array.isArray(result) && Array.isArray(result[0])) {
      return (result as CellValueType[][])[0][0];
    }
    return result as CellValueType;
  }

  // Arithmetic
  it('evaluates basic arithmetic', () => {
    expect(evaluate('1+2')).toBe(3);
    expect(evaluate('10-3')).toBe(7);
    expect(evaluate('4*5')).toBe(20);
    expect(evaluate('15/3')).toBe(5);
    expect(evaluate('2^3')).toBe(8);
  });

  it('handles operator precedence', () => {
    expect(evaluate('1+2*3')).toBe(7);
    expect(evaluate('(1+2)*3')).toBe(9);
  });

  it('handles unary negation', () => {
    expect(evaluate('-5')).toBe(-5);
    expect(evaluate('-5+3')).toBe(-2);
  });

  it('handles division by zero', () => {
    expect(evaluate('1/0')).toBe(FormulaError.DIV0);
  });

  it('handles percentage', () => {
    expect(evaluate('50%')).toBe(0.5);
  });

  // String operations
  it('concatenates strings', () => {
    expect(evaluate('"hello"&" world"')).toBe('hello world');
  });

  // Comparisons
  it('evaluates comparisons', () => {
    expect(evaluate('1=1')).toBe(true);
    expect(evaluate('1<>2')).toBe(true);
    expect(evaluate('1<2')).toBe(true);
    expect(evaluate('2>1')).toBe(true);
    expect(evaluate('1>=1')).toBe(true);
    expect(evaluate('2<=1')).toBe(false);
  });

  // Cell references
  it('resolves cell references', () => {
    expect(evaluate('A1+B1', { A1: 10, B1: 20 })).toBe(30);
  });

  it('returns null for empty cells', () => {
    expect(evaluate('A1', {})).toBe(null);
  });

  // Functions
  it('evaluates SUM', () => {
    expect(evaluate('SUM(A1:C1)', { A1: 1, B1: 2, C1: 3 })).toBe(6);
  });

  it('evaluates AVERAGE', () => {
    expect(evaluate('AVERAGE(A1:C1)', { A1: 10, B1: 20, C1: 30 })).toBe(20);
  });

  it('evaluates IF (true branch)', () => {
    expect(evaluate('IF(A1>5,10,20)', { A1: 10 })).toBe(10);
  });

  it('evaluates IF (false branch)', () => {
    expect(evaluate('IF(A1>5,10,20)', { A1: 1 })).toBe(20);
  });

  it('evaluates nested IF', () => {
    expect(evaluate('IF(A1>10,"big",IF(A1>5,"medium","small"))', { A1: 7 })).toBe('medium');
  });

  it('evaluates IFERROR', () => {
    expect(evaluate('IFERROR(1/0,"error")')).toBe('error');
    expect(evaluate('IFERROR(42,"error")')).toBe(42);
  });

  it('evaluates MIN/MAX', () => {
    expect(evaluate('MIN(A1:C1)', { A1: 5, B1: 1, C1: 10 })).toBe(1);
    expect(evaluate('MAX(A1:C1)', { A1: 5, B1: 1, C1: 10 })).toBe(10);
  });

  it('evaluates COUNT', () => {
    expect(evaluate('COUNT(A1:C1)', { A1: 1, B1: 'text', C1: 3 })).toBe(2);
  });

  it('evaluates CONCATENATE', () => {
    expect(evaluate('CONCATENATE("a","b","c")')).toBe('abc');
  });

  it('evaluates ABS', () => {
    expect(evaluate('ABS(-5)')).toBe(5);
  });

  it('evaluates ROUND', () => {
    expect(evaluate('ROUND(3.14159,2)')).toBeCloseTo(3.14);
  });

  it('returns #NAME? for unknown functions', () => {
    expect(evaluate('UNKNOWNFUNC(1)')).toBe(FormulaError.NAME);
  });

  // VLOOKUP
  it('evaluates VLOOKUP exact match', () => {
    const cells = {
      A1: 'apple',
      B1: 10,
      A2: 'banana',
      B2: 20,
      A3: 'cherry',
      B3: 30,
    };
    expect(evaluate('VLOOKUP("banana",A1:B3,2,FALSE())', cells)).toBe(20);
  });

  // SUMIF
  it('evaluates SUMIF', () => {
    const cells = { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 };
    expect(evaluate('SUMIF(A1:A3,">1",B1:B3)', cells)).toBe(50);
  });

  // COUNTIF
  it('evaluates COUNTIF', () => {
    const cells = { A1: 1, A2: 2, A3: 3 };
    expect(evaluate('COUNTIF(A1:A3,">=2")', cells)).toBe(2);
  });

  // LEN
  it('evaluates LEN', () => {
    expect(evaluate('LEN("hello")')).toBe(5);
  });

  // LEFT/RIGHT/MID
  it('evaluates LEFT', () => {
    expect(evaluate('LEFT("hello",3)')).toBe('hel');
  });

  it('evaluates RIGHT', () => {
    expect(evaluate('RIGHT("hello",3)')).toBe('llo');
  });

  it('evaluates MID', () => {
    expect(evaluate('MID("hello",2,3)')).toBe('ell');
  });

  // AND/OR/NOT
  it('evaluates AND', () => {
    expect(evaluate('AND(TRUE(),TRUE())')).toBe(true);
    expect(evaluate('AND(TRUE(),FALSE())')).toBe(false);
  });

  it('evaluates OR', () => {
    expect(evaluate('OR(FALSE(),TRUE())')).toBe(true);
    expect(evaluate('OR(FALSE(),FALSE())')).toBe(false);
  });

  it('evaluates NOT', () => {
    expect(evaluate('NOT(TRUE())')).toBe(false);
    expect(evaluate('NOT(FALSE())')).toBe(true);
  });
});
