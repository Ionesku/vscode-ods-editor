import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../../src/formula/Tokenizer';
import { Parser } from '../../../../src/formula/Parser';
import { Evaluator, CellValueGetter } from '../../../../src/formula/Evaluator';
import { createDefaultRegistry } from '../../../../src/formula/functions/index';
import { FormulaError } from '../../../../src/model/types';
import { CellValueType } from '../../../../src/formula/types';

describe('Math Functions', () => {
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

  describe('SUM', () => {
    it('sums numbers', () => {
      expect(evaluate('SUM(A1:C1)', { A1: 1, B1: 2, C1: 3 })).toBe(6);
    });
    it('ignores strings', () => {
      expect(evaluate('SUM(A1:C1)', { A1: 1, B1: 'text', C1: 3 })).toBe(4);
    });
    it('returns 0 for empty range', () => {
      expect(evaluate('SUM(A1:C1)', {})).toBe(0);
    });
    it('handles multiple arguments', () => {
      expect(evaluate('SUM(1,2,3)')).toBe(6);
    });
  });

  describe('AVERAGE', () => {
    it('computes average', () => {
      expect(evaluate('AVERAGE(A1:C1)', { A1: 10, B1: 20, C1: 30 })).toBe(20);
    });
    it('returns DIV0 for empty range', () => {
      expect(evaluate('AVERAGE(A1:C1)', {})).toBe(FormulaError.DIV0);
    });
  });

  describe('MIN/MAX', () => {
    it('finds minimum', () => {
      expect(evaluate('MIN(A1:C1)', { A1: 5, B1: 1, C1: 10 })).toBe(1);
    });
    it('finds maximum', () => {
      expect(evaluate('MAX(A1:C1)', { A1: 5, B1: 1, C1: 10 })).toBe(10);
    });
    it('returns 0 for no numbers', () => {
      expect(evaluate('MIN(A1:C1)', {})).toBe(0);
    });
  });

  describe('ROUND', () => {
    it('rounds to specified digits', () => {
      expect(evaluate('ROUND(3.14159,2)')).toBeCloseTo(3.14);
      expect(evaluate('ROUND(3.145,2)')).toBeCloseTo(3.15);
      expect(evaluate('ROUND(123,-1)')).toBe(120);
    });
  });

  describe('ABS', () => {
    it('returns absolute value', () => {
      expect(evaluate('ABS(-5)')).toBe(5);
      expect(evaluate('ABS(5)')).toBe(5);
      expect(evaluate('ABS(0)')).toBe(0);
    });
  });

  describe('MOD', () => {
    it('computes modulo', () => {
      expect(evaluate('MOD(10,3)')).toBe(1);
      expect(evaluate('MOD(10,5)')).toBe(0);
    });
    it('returns DIV0 for zero divisor', () => {
      expect(evaluate('MOD(10,0)')).toBe(FormulaError.DIV0);
    });
  });

  describe('POWER', () => {
    it('computes power', () => {
      expect(evaluate('POWER(2,10)')).toBe(1024);
    });
  });

  describe('SQRT', () => {
    it('computes square root', () => {
      expect(evaluate('SQRT(25)')).toBe(5);
    });
    it('returns NUM for negative', () => {
      expect(evaluate('SQRT(-1)')).toBe(FormulaError.NUM);
    });
  });

  describe('INT', () => {
    it('floors to integer', () => {
      expect(evaluate('INT(3.7)')).toBe(3);
      expect(evaluate('INT(-3.7)')).toBe(-4);
    });
  });

  describe('SUMIF', () => {
    const cells = { A1: 10, A2: 20, A3: 30, B1: 1, B2: 2, B3: 3 };
    it('sums with criteria', () => {
      expect(evaluate('SUMIF(A1:A3,">15",B1:B3)', cells)).toBe(5);
    });
    it('sums with exact match', () => {
      expect(evaluate('SUMIF(A1:A3,20,B1:B3)', cells)).toBe(2);
    });
  });

  describe('PI', () => {
    it('returns pi', () => {
      expect(evaluate('PI()')).toBeCloseTo(Math.PI);
    });
  });

  describe('PRODUCT', () => {
    it('multiplies numbers', () => {
      expect(evaluate('PRODUCT(A1:C1)', { A1: 2, B1: 3, C1: 4 })).toBe(24);
    });
  });

  describe('SUMPRODUCT', () => {
    it('computes sum of products', () => {
      expect(
        evaluate('SUMPRODUCT(A1:A3,B1:B3)', {
          A1: 1,
          A2: 2,
          A3: 3,
          B1: 4,
          B2: 5,
          B3: 6,
        }),
      ).toBe(32); // 1*4 + 2*5 + 3*6
    });
  });
});
