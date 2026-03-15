import { describe, it, expect } from 'vitest';
import { FormulaEngine } from '../../../src/formula/FormulaEngine';
import { SpreadsheetModel } from '../../../src/model/SpreadsheetModel';
import { FormulaError } from '../../../src/model/types';

describe('FormulaEngine', () => {
  function makeEngine() {
    return new FormulaEngine();
  }

  function makeModel() {
    return new SpreadsheetModel();
  }

  describe('recalcAll - basic formulas', () => {
    it('evaluates a simple constant formula', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: null, formula: '1+2', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(0, 0).computedValue).toBe(3);
    });

    it('evaluates a cell reference formula', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 10, computedValue: 10 });
      model.sheets[0].setCell(1, 0, { rawValue: null, formula: 'A1*2', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(1, 0).computedValue).toBe(20);
    });

    it('evaluates SUM over a range', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 1, computedValue: 1 });
      model.sheets[0].setCell(0, 1, { rawValue: 2, computedValue: 2 });
      model.sheets[0].setCell(0, 2, { rawValue: 3, computedValue: 3 });
      model.sheets[0].setCell(0, 3, { rawValue: null, formula: 'SUM(A1:A3)', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(0, 3).computedValue).toBe(6);
    });

    it('evaluates IF formula', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 10, computedValue: 10 });
      model.sheets[0].setCell(1, 0, {
        rawValue: null,
        formula: 'IF(A1>5,"big","small")',
        computedValue: null,
      });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(1, 0).computedValue).toBe('big');
    });
  });

  describe('recalcAll - dependency ordering', () => {
    it('evaluates chain: A1 → B1 → C1', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 5, computedValue: 5 });
      model.sheets[0].setCell(1, 0, { rawValue: null, formula: 'A1+1', computedValue: null });
      model.sheets[0].setCell(2, 0, { rawValue: null, formula: 'B1*2', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(1, 0).computedValue).toBe(6);
      expect(model.sheets[0].getCell(2, 0).computedValue).toBe(12);
    });

    it('recalculates diamond dependency correctly', () => {
      // A1 = 3
      // B1 = A1 + 1  (4)
      // C1 = A1 * 2  (6)
      // D1 = B1 + C1 (10)
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 3, computedValue: 3 });
      model.sheets[0].setCell(1, 0, { rawValue: null, formula: 'A1+1', computedValue: null });
      model.sheets[0].setCell(2, 0, { rawValue: null, formula: 'A1*2', computedValue: null });
      model.sheets[0].setCell(3, 0, { rawValue: null, formula: 'B1+C1', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(3, 0).computedValue).toBe(10);
    });

    it('recalculating multiple times gives consistent results', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 7, computedValue: 7 });
      model.sheets[0].setCell(1, 0, { rawValue: null, formula: 'A1+3', computedValue: null });

      engine.recalcAll(model);
      engine.recalcAll(model);

      expect(model.sheets[0].getCell(1, 0).computedValue).toBe(10);
    });
  });

  describe('recalcAll - error handling', () => {
    it('marks invalid formula with #NAME?', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, {
        rawValue: null,
        formula: 'NONEXISTENTFUNC(1)',
        computedValue: null,
      });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(0, 0).computedValue).toBe(FormulaError.NAME);
    });

    it('propagates #DIV/0! error', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: null, formula: '1/0', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(0, 0).computedValue).toBe(FormulaError.DIV0);
    });

    it('handles reference to missing sheet gracefully', () => {
      const engine = makeEngine();
      const model = makeModel();
      // A formula referencing a nonexistent sheet
      model.sheets[0].setCell(0, 0, { rawValue: null, formula: 'SUM(A1:A3)', computedValue: null });

      // Should not throw
      expect(() => engine.recalcAll(model)).not.toThrow();
    });
  });

  describe('recalcAll - cross-sheet references', () => {
    it('evaluates formula referencing another sheet', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.addSheet('Sheet2');
      model.sheets[1].setCell(0, 0, { rawValue: 42, computedValue: 42 });
      // Reference Sheet2.A1 from Sheet1
      model.sheets[0].setCell(0, 0, { rawValue: null, formula: 'Sheet2.A1', computedValue: null });

      engine.recalcAll(model);

      // Sheet2.A1 = 42, so Sheet1.A1 should be 42
      expect(model.sheets[0].getCell(0, 0).computedValue).toBe(42);
    });
  });

  describe('recalcAll - ODS formula prefix', () => {
    it('handles of:= prefix in stored formulas', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: null, formula: 'of:=2+3', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(0, 0).computedValue).toBe(5);
    });
  });

  describe('recalcAll - multiple sheets', () => {
    it('recalculates formulas on all sheets', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.addSheet('Sheet2');
      model.sheets[0].setCell(0, 0, { rawValue: null, formula: '10+5', computedValue: null });
      model.sheets[1].setCell(0, 0, { rawValue: null, formula: '3*4', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(0, 0).computedValue).toBe(15);
      expect(model.sheets[1].getCell(0, 0).computedValue).toBe(12);
    });
  });

  describe('recalcAll - named ranges', () => {
    it('evaluates formula using a named range', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 100, computedValue: 100 });
      model.sheets[0].setCell(1, 0, { rawValue: 200, computedValue: 200 });
      model.defineNamedRange('Total', 'Sheet1', {
        sheet: 'Sheet1',
        startCol: 0,
        startRow: 0,
        endCol: 1,
        endRow: 0,
      });
      model.sheets[0].setCell(0, 1, { rawValue: null, formula: 'SUM(Total)', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(0, 1).computedValue).toBe(300);
    });
  });

  describe('array formulas (CSE)', () => {
    it('element-wise multiply then SUM — {=SUM(A1:A3*B1:B3)}', () => {
      const engine = makeEngine();
      const model = makeModel();
      // A1:A3 = [1,2,3], B1:B3 = [10,20,30]
      model.sheets[0].setCell(0, 0, { rawValue: 1, computedValue: 1 });
      model.sheets[0].setCell(0, 1, { rawValue: 2, computedValue: 2 });
      model.sheets[0].setCell(0, 2, { rawValue: 3, computedValue: 3 });
      model.sheets[0].setCell(1, 0, { rawValue: 10, computedValue: 10 });
      model.sheets[0].setCell(1, 1, { rawValue: 20, computedValue: 20 });
      model.sheets[0].setCell(1, 2, { rawValue: 30, computedValue: 30 });
      // {=SUM(A1:A3*B1:B3)} = 1*10 + 2*20 + 3*30 = 10+40+90 = 140
      model.sheets[0].setCell(2, 0, {
        rawValue: null,
        formula: '{=SUM(A1:A3*B1:B3)}',
        computedValue: null,
      });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(2, 0).computedValue).toBe(140);
    });

    it('element-wise addition of two ranges', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 1, computedValue: 1 });
      model.sheets[0].setCell(0, 1, { rawValue: 2, computedValue: 2 });
      model.sheets[0].setCell(1, 0, { rawValue: 10, computedValue: 10 });
      model.sheets[0].setCell(1, 1, { rawValue: 20, computedValue: 20 });
      // SUM(A1:A2 + B1:B2) = (1+10) + (2+20) = 33
      model.sheets[0].setCell(2, 0, {
        rawValue: null,
        formula: 'SUM(A1:A2+B1:B2)',
        computedValue: null,
      });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(2, 0).computedValue).toBe(33);
    });

    it('range multiplied by a scalar', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 5, computedValue: 5 });
      model.sheets[0].setCell(0, 1, { rawValue: 6, computedValue: 6 });
      // SUM(A1:A2*3) = 15+18 = 33
      model.sheets[0].setCell(1, 0, {
        rawValue: null,
        formula: 'SUM(A1:A2*3)',
        computedValue: null,
      });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(1, 0).computedValue).toBe(33);
    });

    it('strips {= ... } braces and evaluates normally', () => {
      const engine = makeEngine();
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 7, computedValue: 7 });
      model.sheets[0].setCell(1, 0, { rawValue: null, formula: '{=A1*2}', computedValue: null });

      engine.recalcAll(model);

      expect(model.sheets[0].getCell(1, 0).computedValue).toBe(14);
    });
  });
});
