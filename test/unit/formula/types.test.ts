import { describe, it, expect } from 'vitest';
import { isFormulaError, isRangeValue, flattenRange, toNumber } from '../../../src/formula/types';
import { FormulaError } from '../../../src/model/types';

describe('formula/types utilities', () => {
  describe('isFormulaError', () => {
    it('recognises all FormulaError enum values', () => {
      expect(isFormulaError(FormulaError.REF)).toBe(true);
      expect(isFormulaError(FormulaError.VALUE)).toBe(true);
      expect(isFormulaError(FormulaError.DIV0)).toBe(true);
      expect(isFormulaError(FormulaError.NAME)).toBe(true);
      expect(isFormulaError(FormulaError.NA)).toBe(true);
      expect(isFormulaError(FormulaError.NUM)).toBe(true);
      expect(isFormulaError(FormulaError.NULL)).toBe(true);
    });

    it('returns false for normal values', () => {
      expect(isFormulaError(42)).toBe(false);
      expect(isFormulaError('hello')).toBe(false);
      expect(isFormulaError(true)).toBe(false);
      expect(isFormulaError(null)).toBe(false);
    });

    it('returns false for strings that look error-like but are not', () => {
      expect(isFormulaError('#notanerror')).toBe(false);
      expect(isFormulaError('')).toBe(false);
    });
  });

  describe('isRangeValue', () => {
    it('returns true for 2D arrays', () => {
      expect(
        isRangeValue([
          [1, 2],
          [3, 4],
        ]),
      ).toBe(true);
      expect(isRangeValue([[]])).toBe(true);
    });

    it('returns true for empty array (empty range)', () => {
      expect(isRangeValue([])).toBe(true);
    });

    it('returns false for scalars', () => {
      expect(isRangeValue(42)).toBe(false);
      expect(isRangeValue('hello')).toBe(false);
      expect(isRangeValue(true)).toBe(false);
      expect(isRangeValue(null)).toBe(false);
    });

    it('returns false for flat arrays (not 2D)', () => {
      // A flat array is not a range — first element is not an array
      expect(
        isRangeValue([1, 2, 3] as unknown as import('../../../src/formula/types').FormulaValue),
      ).toBe(false);
    });
  });

  describe('flattenRange', () => {
    it('flattens 2D range to 1D array', () => {
      const result = flattenRange([
        [1, 2],
        [3, 4],
      ]);
      expect(result).toEqual([1, 2, 3, 4]);
    });

    it('wraps scalar in array', () => {
      expect(flattenRange(42)).toEqual([42]);
      expect(flattenRange('hello')).toEqual(['hello']);
      expect(flattenRange(null)).toEqual([null]);
    });

    it('flattens single-row range', () => {
      expect(flattenRange([[10, 20, 30]])).toEqual([10, 20, 30]);
    });

    it('flattens single-column range', () => {
      expect(flattenRange([[10], [20], [30]])).toEqual([10, 20, 30]);
    });

    it('flattens empty range', () => {
      expect(flattenRange([])).toEqual([]);
    });

    it('preserves nulls and errors in range', () => {
      const result = flattenRange([
        [1, null],
        [FormulaError.DIV0, 'text'],
      ]);
      expect(result).toEqual([1, null, FormulaError.DIV0, 'text']);
    });
  });

  describe('toNumber', () => {
    it('returns number as-is', () => {
      expect(toNumber(42)).toBe(42);
      expect(toNumber(0)).toBe(0);
      expect(toNumber(-3.14)).toBe(-3.14);
    });

    it('converts boolean true to 1', () => {
      expect(toNumber(true)).toBe(1);
    });

    it('converts boolean false to 0', () => {
      expect(toNumber(false)).toBe(0);
    });

    it('converts null to 0', () => {
      expect(toNumber(null)).toBe(0);
    });

    it('converts numeric string to number', () => {
      expect(toNumber('42')).toBe(42);
      expect(toNumber('3.14')).toBeCloseTo(3.14);
      expect(toNumber('0')).toBe(0);
      expect(toNumber('-5')).toBe(-5);
    });

    it('converts empty string to 0', () => {
      expect(toNumber('')).toBe(0);
    });

    it('returns #VALUE! for non-numeric string', () => {
      expect(toNumber('hello')).toBe(FormulaError.VALUE);
      expect(toNumber('12abc')).toBe(FormulaError.VALUE);
    });

    it('returns #VALUE! for formula error strings (they are unparseable strings)', () => {
      // FormulaError values are string enums like '#DIV/0!', '#REF!' — they fail Number() parsing
      expect(toNumber(FormulaError.DIV0)).toBe(FormulaError.VALUE);
      expect(toNumber(FormulaError.REF)).toBe(FormulaError.VALUE);
      expect(toNumber(FormulaError.NUM)).toBe(FormulaError.VALUE);
    });
  });
});
