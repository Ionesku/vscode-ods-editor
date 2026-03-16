import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../../src/formula/Tokenizer';
import { Parser } from '../../../../src/formula/Parser';
import { Evaluator, CellValueGetter } from '../../../../src/formula/Evaluator';
import { createDefaultRegistry } from '../../../../src/formula/functions/index';
import { FormulaError } from '../../../../src/model/types';
import { CellValueType } from '../../../../src/formula/types';

describe('Statistical Functions', () => {
  const tokenizer = new Tokenizer();
  const parser = new Parser();
  const registry = createDefaultRegistry();

  function evaluate(formula: string, cells: Record<string, CellValueType> = {}): CellValueType {
    const getCellValue: CellValueGetter = (sheet, col, row) => {
      const key = `${String.fromCharCode(65 + col)}${row + 1}`;
      return cells[key] ?? null;
    };
    const evaluator = new Evaluator(getCellValue, registry);
    const ast = parser.parse(tokenizer.tokenize(formula));
    const result = evaluator.evaluate(ast, 'Sheet1');
    if (Array.isArray(result) && Array.isArray(result[0]))
      return (result as CellValueType[][])[0][0];
    return result as CellValueType;
  }

  it('COUNT counts numbers', () => {
    expect(evaluate('COUNT(A1:C1)', { A1: 1, B1: 'text', C1: 3 })).toBe(2);
  });

  it('COUNTA counts non-empty', () => {
    expect(evaluate('COUNTA(A1:C1)', { A1: 1, B1: 'text', C1: 3 })).toBe(3);
  });

  it('COUNTBLANK counts empty', () => {
    expect(evaluate('COUNTBLANK(A1:C1)', { A1: 1 })).toBe(2);
  });

  it('COUNTIF with criteria', () => {
    const cells = { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5 };
    expect(evaluate('COUNTIF(A1:A5,">=3")', cells)).toBe(3);
  });

  it('MEDIAN', () => {
    expect(evaluate('MEDIAN(A1:E1)', { A1: 1, B1: 3, C1: 5, D1: 7, E1: 9 })).toBe(5);
    expect(evaluate('MEDIAN(A1:D1)', { A1: 1, B1: 3, C1: 5, D1: 7 })).toBe(4);
  });

  it('STDEV', () => {
    const cells = { A1: 2, A2: 4, A3: 4, A4: 4, A5: 5, A6: 5, A7: 7, A8: 9 };
    const result = evaluate('STDEV(A1:A8)', cells);
    expect(typeof result).toBe('number');
    expect(result as number).toBeCloseTo(2.1381, 3);
  });

  it('LARGE', () => {
    expect(evaluate('LARGE(A1:E1,2)', { A1: 10, B1: 30, C1: 20, D1: 50, E1: 40 })).toBe(40);
  });

  it('SMALL', () => {
    expect(evaluate('SMALL(A1:E1,2)', { A1: 10, B1: 30, C1: 20, D1: 50, E1: 40 })).toBe(20);
  });

  it('PERCENTILE', () => {
    const cells = { A1: 1, A2: 2, A3: 3, A4: 4, A5: 5 };
    expect(evaluate('PERCENTILE(A1:A5,0.5)', cells)).toBe(3);
  });

  it('RANK', () => {
    const cells = { A1: 10, A2: 30, A3: 20 };
    expect(evaluate('RANK(30,A1:A3)', cells)).toBe(1); // Descending
  });
});
