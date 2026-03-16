import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../../src/formula/Tokenizer';
import { Parser } from '../../../../src/formula/Parser';
import { Evaluator, CellValueGetter } from '../../../../src/formula/Evaluator';
import { createDefaultRegistry } from '../../../../src/formula/functions/index';
import { FormulaError } from '../../../../src/model/types';
import { CellValueType } from '../../../../src/formula/types';

describe('Lookup Functions', () => {
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

  describe('VLOOKUP', () => {
    const cells = {
      A1: 'apple',
      B1: 1,
      A2: 'banana',
      B2: 2,
      A3: 'cherry',
      B3: 3,
    };

    it('exact match found', () => {
      expect(evaluate('VLOOKUP("banana",A1:B3,2,FALSE())', cells)).toBe(2);
    });

    it('exact match not found', () => {
      expect(evaluate('VLOOKUP("grape",A1:B3,2,FALSE())', cells)).toBe(FormulaError.NA);
    });

    it('case-insensitive match', () => {
      expect(evaluate('VLOOKUP("BANANA",A1:B3,2,FALSE())', cells)).toBe(2);
    });
  });

  describe('INDEX', () => {
    const cells = {
      A1: 1,
      B1: 2,
      C1: 3,
      A2: 4,
      B2: 5,
      C2: 6,
      A3: 7,
      B3: 8,
      C3: 9,
    };

    it('returns correct cell', () => {
      expect(evaluate('INDEX(A1:C3,2,3)', cells)).toBe(6);
    });

    it('out of range returns REF', () => {
      expect(evaluate('INDEX(A1:C3,5,1)', cells)).toBe(FormulaError.REF);
    });
  });

  describe('MATCH', () => {
    const cells = { A1: 10, A2: 20, A3: 30 };

    it('exact match', () => {
      expect(evaluate('MATCH(20,A1:A3,0)', cells)).toBe(2);
    });

    it('not found', () => {
      expect(evaluate('MATCH(25,A1:A3,0)', cells)).toBe(FormulaError.NA);
    });
  });

  describe('ROWS/COLUMNS', () => {
    it('counts rows', () => {
      expect(evaluate('ROWS(A1:A5)', {})).toBe(5);
    });
    it('counts columns', () => {
      expect(evaluate('COLUMNS(A1:C1)', {})).toBe(3);
    });
  });
});
