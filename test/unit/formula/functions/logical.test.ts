import { describe, it, expect } from 'vitest';
import { Tokenizer } from '../../../../src/formula/Tokenizer';
import { Parser } from '../../../../src/formula/Parser';
import { Evaluator, CellValueGetter } from '../../../../src/formula/Evaluator';
import { createDefaultRegistry } from '../../../../src/formula/functions/index';
import { CellValueType } from '../../../../src/formula/types';

describe('Logical Functions', () => {
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

  describe('IF', () => {
    it('returns true branch when condition is true', () => {
      expect(evaluate('IF(TRUE(),"yes","no")')).toBe('yes');
    });
    it('returns false branch when condition is false', () => {
      expect(evaluate('IF(FALSE(),"yes","no")')).toBe('no');
    });
    it('returns false when no else branch and condition is false', () => {
      expect(evaluate('IF(FALSE(),"yes")')).toBe(false);
    });
    it('evaluates numeric conditions', () => {
      expect(evaluate('IF(1,"yes","no")')).toBe('yes');
      expect(evaluate('IF(0,"yes","no")')).toBe('no');
    });
  });

  describe('AND', () => {
    it('returns true when all true', () => {
      expect(evaluate('AND(TRUE(),TRUE(),TRUE())')).toBe(true);
    });
    it('returns false when any false', () => {
      expect(evaluate('AND(TRUE(),FALSE(),TRUE())')).toBe(false);
    });
    it('works with numbers', () => {
      expect(evaluate('AND(1,2,3)')).toBe(true);
      expect(evaluate('AND(1,0,3)')).toBe(false);
    });
  });

  describe('OR', () => {
    it('returns true when any true', () => {
      expect(evaluate('OR(FALSE(),TRUE(),FALSE())')).toBe(true);
    });
    it('returns false when all false', () => {
      expect(evaluate('OR(FALSE(),FALSE())')).toBe(false);
    });
  });

  describe('NOT', () => {
    it('negates boolean', () => {
      expect(evaluate('NOT(TRUE())')).toBe(false);
      expect(evaluate('NOT(FALSE())')).toBe(true);
    });
  });

  describe('ISBLANK', () => {
    it('detects blank cells', () => {
      expect(evaluate('ISBLANK(A1)', {})).toBe(true);
      expect(evaluate('ISBLANK(A1)', { A1: 'text' })).toBe(false);
    });
  });

  describe('ISNUMBER', () => {
    it('detects numbers', () => {
      expect(evaluate('ISNUMBER(A1)', { A1: 42 })).toBe(true);
      expect(evaluate('ISNUMBER(A1)', { A1: 'text' })).toBe(false);
    });
  });
});
