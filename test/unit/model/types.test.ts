import { describe, it, expect } from 'vitest';
import {
  colToLetter,
  letterToCol,
  cellKey,
  parseCellKey,
  createEmptyCell,
} from '../../../src/model/types';

describe('model/types utilities', () => {
  describe('colToLetter', () => {
    it('converts single-letter columns', () => {
      expect(colToLetter(0)).toBe('A');
      expect(colToLetter(1)).toBe('B');
      expect(colToLetter(25)).toBe('Z');
    });

    it('converts double-letter columns', () => {
      expect(colToLetter(26)).toBe('AA');
      expect(colToLetter(27)).toBe('AB');
      expect(colToLetter(51)).toBe('AZ');
      expect(colToLetter(52)).toBe('BA');
    });

    it('converts triple-letter columns', () => {
      expect(colToLetter(702)).toBe('AAA');
    });
  });

  describe('letterToCol', () => {
    it('converts single letters', () => {
      expect(letterToCol('A')).toBe(0);
      expect(letterToCol('B')).toBe(1);
      expect(letterToCol('Z')).toBe(25);
    });

    it('converts double letters', () => {
      expect(letterToCol('AA')).toBe(26);
      expect(letterToCol('AB')).toBe(27);
      expect(letterToCol('AZ')).toBe(51);
      expect(letterToCol('BA')).toBe(52);
    });
  });

  describe('colToLetter / letterToCol round-trip', () => {
    it('round-trips for single-letter columns', () => {
      for (let i = 0; i <= 25; i++) {
        expect(letterToCol(colToLetter(i))).toBe(i);
      }
    });

    it('round-trips for double-letter columns', () => {
      for (let i = 26; i <= 100; i++) {
        expect(letterToCol(colToLetter(i))).toBe(i);
      }
    });
  });

  describe('cellKey', () => {
    it('formats key as "col,row"', () => {
      expect(cellKey(0, 0)).toBe('0,0');
      expect(cellKey(3, 7)).toBe('3,7');
      expect(cellKey(100, 999)).toBe('100,999');
    });

    it('produces unique keys for different positions', () => {
      const keys = new Set([cellKey(0, 1), cellKey(1, 0), cellKey(0, 0), cellKey(1, 1)]);
      expect(keys.size).toBe(4);
    });
  });

  describe('parseCellKey', () => {
    it('parses "0,0"', () => {
      expect(parseCellKey('0,0')).toEqual({ col: 0, row: 0 });
    });

    it('parses multi-digit values', () => {
      expect(parseCellKey('100,999')).toEqual({ col: 100, row: 999 });
    });

    it('round-trips with cellKey', () => {
      const positions = [
        { col: 0, row: 0 },
        { col: 5, row: 12 },
        { col: 1023, row: 65535 },
      ];
      for (const pos of positions) {
        expect(parseCellKey(cellKey(pos.col, pos.row))).toEqual(pos);
      }
    });
  });

  describe('createEmptyCell', () => {
    it('creates a cell with all null/default values', () => {
      const cell = createEmptyCell();
      expect(cell.rawValue).toBeNull();
      expect(cell.formula).toBeNull();
      expect(cell.computedValue).toBeNull();
      expect(cell.styleId).toBeNull();
      expect(cell.mergeColSpan).toBe(1);
      expect(cell.mergeRowSpan).toBe(1);
      expect(cell.mergedInto).toBeNull();
    });

    it('creates a fresh object each call (no shared reference)', () => {
      const a = createEmptyCell();
      const b = createEmptyCell();
      a.rawValue = 42;
      expect(b.rawValue).toBeNull();
    });
  });
});
