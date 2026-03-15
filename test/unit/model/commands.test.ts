import { describe, it, expect } from 'vitest';
import { SpreadsheetModel } from '../../../src/model/SpreadsheetModel';
import {
  SetCellValueCommand,
  SetCellStyleCommand,
  MergeCellsCommand,
  InsertRowsCommand,
  InsertColumnsCommand,
  DeleteRowsCommand,
  DeleteColumnsCommand,
  ResizeColumnCommand,
  ResizeRowCommand,
  SortRangeCommand,
  SortRangeMultiCommand,
  PasteSpecialCommand,
} from '../../../src/model/commands';

describe('Commands', () => {
  describe('SetCellValueCommand', () => {
    it('sets and undoes cell value', () => {
      const model = new SpreadsheetModel();
      const cmd = new SetCellValueCommand('Sheet1', 0, 0, 42, null);
      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBe(42);

      cmd.undo(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBeNull();
    });

    it('sets formula', () => {
      const model = new SpreadsheetModel();
      const cmd = new SetCellValueCommand('Sheet1', 0, 0, null, 'A2+B2');
      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).formula).toBe('A2+B2');
    });

    it('old value is captured only on first execute', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'original', computedValue: 'original' });
      const cmd = new SetCellValueCommand('Sheet1', 0, 0, 'first', null);
      cmd.execute(model);

      // Execute again — old value should still be 'original'
      const cmd2 = new SetCellValueCommand('Sheet1', 0, 0, 'second', null);
      cmd2.execute(model);
      cmd.execute(model); // re-execute cmd, but old is already captured
      cmd.undo(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBe('original');
    });

    it('does nothing when sheet does not exist', () => {
      const model = new SpreadsheetModel();
      const cmd = new SetCellValueCommand('NoSheet', 0, 0, 42, null);
      expect(() => cmd.execute(model)).not.toThrow();
      expect(() => cmd.undo(model)).not.toThrow();
    });
  });

  describe('SetCellStyleCommand', () => {
    it('sets and undoes style', () => {
      const model = new SpreadsheetModel();
      model.styles.set('bold', { id: 'bold', bold: true });
      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 1, endRow: 1 };
      const cmd = new SetCellStyleCommand('Sheet1', range, 'bold');

      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).styleId).toBe('bold');
      expect(model.sheets[0].getCell(1, 1).styleId).toBe('bold');

      cmd.undo(model);
      expect(model.sheets[0].getCell(0, 0).styleId).toBeNull();
    });

    it('restores previous style on undo', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { styleId: 'old-style' });
      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 0 };
      const cmd = new SetCellStyleCommand('Sheet1', range, 'new-style');
      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).styleId).toBe('new-style');
      cmd.undo(model);
      expect(model.sheets[0].getCell(0, 0).styleId).toBe('old-style');
    });
  });

  describe('MergeCellsCommand', () => {
    it('merges and unmerges cells', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'merged', computedValue: 'merged' });
      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 2, endRow: 1 };
      const cmd = new MergeCellsCommand('Sheet1', range);

      cmd.execute(model);
      const anchor = model.sheets[0].getCell(0, 0);
      expect(anchor.mergeColSpan).toBe(3);
      expect(anchor.mergeRowSpan).toBe(2);
      const covered = model.sheets[0].getCell(1, 0);
      expect(covered.mergedInto).not.toBeNull();

      cmd.undo(model);
      const unmerged = model.sheets[0].getCell(0, 0);
      expect(unmerged.mergeColSpan).toBe(1);
    });

    it('covered cells lose their values', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'anchor', computedValue: 'anchor' });
      model.sheets[0].setCell(1, 0, { rawValue: 'covered', computedValue: 'covered' });
      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 1, endRow: 0 };
      const cmd = new MergeCellsCommand('Sheet1', range);

      cmd.execute(model);
      expect(model.sheets[0].getCell(1, 0).rawValue).toBeNull();
      expect(model.sheets[0].getCell(1, 0).mergedInto).not.toBeNull();

      // Undo restores covered cell value
      cmd.undo(model);
      expect(model.sheets[0].getCell(1, 0).rawValue).toBe('covered');
    });
  });

  describe('InsertRowsCommand', () => {
    it('inserts and undoes rows', () => {
      const model = new SpreadsheetModel();
      const origRowCount = model.sheets[0].rowCount;
      const cmd = new InsertRowsCommand('Sheet1', 5, 3);

      cmd.execute(model);
      expect(model.sheets[0].rowCount).toBe(origRowCount + 3);

      cmd.undo(model);
      expect(model.sheets[0].rowCount).toBe(origRowCount);
    });

    it('shifts cells down on insert', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 2, { rawValue: 'below', computedValue: 'below' });
      const cmd = new InsertRowsCommand('Sheet1', 1, 3);
      cmd.execute(model);
      // Cell that was at row 2 should now be at row 5
      expect(model.sheets[0].getCell(0, 5).rawValue).toBe('below');
      expect(model.sheets[0].getCell(0, 2).rawValue).toBeNull();
    });
  });

  describe('InsertColumnsCommand', () => {
    it('inserts and undoes columns', () => {
      const model = new SpreadsheetModel();
      const origColCount = model.sheets[0].columnCount;
      const cmd = new InsertColumnsCommand('Sheet1', 2, 2);

      cmd.execute(model);
      expect(model.sheets[0].columnCount).toBe(origColCount + 2);

      cmd.undo(model);
      expect(model.sheets[0].columnCount).toBe(origColCount);
    });

    it('shifts cells right on insert', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(3, 0, { rawValue: 'right', computedValue: 'right' });
      const cmd = new InsertColumnsCommand('Sheet1', 1, 2);
      cmd.execute(model);
      // Column 3 shifts to column 5
      expect(model.sheets[0].getCell(5, 0).rawValue).toBe('right');
      expect(model.sheets[0].getCell(3, 0).rawValue).toBeNull();
    });
  });

  describe('DeleteRowsCommand', () => {
    it('deletes rows and shifts remaining up', () => {
      const model = new SpreadsheetModel();
      const origRowCount = model.sheets[0].rowCount;
      model.sheets[0].setCell(0, 0, { rawValue: 'top', computedValue: 'top' });
      model.sheets[0].setCell(0, 2, { rawValue: 'bottom', computedValue: 'bottom' });
      const cmd = new DeleteRowsCommand('Sheet1', 1, 1);

      cmd.execute(model);
      expect(model.sheets[0].rowCount).toBe(origRowCount - 1);
      expect(model.sheets[0].getCell(0, 1).rawValue).toBe('bottom');
    });

    it('undoes deletion and restores cells', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 1, { rawValue: 'deleted', computedValue: 'deleted' });
      const cmd = new DeleteRowsCommand('Sheet1', 1, 1);
      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 1).rawValue).toBeNull();

      cmd.undo(model);
      expect(model.sheets[0].getCell(0, 1).rawValue).toBe('deleted');
    });
  });

  describe('DeleteColumnsCommand', () => {
    it('deletes columns and shifts remaining left', () => {
      const model = new SpreadsheetModel();
      const origColCount = model.sheets[0].columnCount;
      model.sheets[0].setCell(3, 0, { rawValue: 'right', computedValue: 'right' });
      const cmd = new DeleteColumnsCommand('Sheet1', 1, 2);

      cmd.execute(model);
      expect(model.sheets[0].columnCount).toBe(origColCount - 2);
      expect(model.sheets[0].getCell(1, 0).rawValue).toBe('right');
    });

    it('undoes deletion and restores cells', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(2, 0, { rawValue: 'deleted', computedValue: 'deleted' });
      const cmd = new DeleteColumnsCommand('Sheet1', 2, 1);
      cmd.execute(model);
      expect(model.sheets[0].getCell(2, 0).rawValue).toBeNull();

      cmd.undo(model);
      expect(model.sheets[0].getCell(2, 0).rawValue).toBe('deleted');
    });
  });

  describe('ResizeColumnCommand', () => {
    it('resizes and undoes column width', () => {
      const model = new SpreadsheetModel();
      const cmd = new ResizeColumnCommand('Sheet1', 0, 150);

      cmd.execute(model);
      expect(model.sheets[0].getColumnWidth(0)).toBe(150);

      cmd.undo(model);
      expect(model.sheets[0].getColumnWidth(0)).toBe(80);
    });
  });

  describe('ResizeRowCommand', () => {
    it('resizes and undoes row height', () => {
      const model = new SpreadsheetModel();
      const cmd = new ResizeRowCommand('Sheet1', 0, 60);

      cmd.execute(model);
      expect(model.sheets[0].getRowHeight(0)).toBe(60);

      cmd.undo(model);
      expect(model.sheets[0].getRowHeight(0)).toBe(24); // default
    });
  });

  describe('SortRangeCommand', () => {
    it('sorts ascending and undoes', () => {
      const model = new SpreadsheetModel();
      const sheet = model.sheets[0];
      sheet.setCell(0, 0, { rawValue: 3, computedValue: 3 });
      sheet.setCell(0, 1, { rawValue: 1, computedValue: 1 });
      sheet.setCell(0, 2, { rawValue: 2, computedValue: 2 });

      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 2 };
      const cmd = new SortRangeCommand('Sheet1', range, 0, true);

      cmd.execute(model);
      expect(sheet.getCell(0, 0).rawValue).toBe(1);
      expect(sheet.getCell(0, 1).rawValue).toBe(2);
      expect(sheet.getCell(0, 2).rawValue).toBe(3);

      cmd.undo(model);
      expect(sheet.getCell(0, 0).rawValue).toBe(3);
      expect(sheet.getCell(0, 1).rawValue).toBe(1);
      expect(sheet.getCell(0, 2).rawValue).toBe(2);
    });

    it('sorts descending', () => {
      const model = new SpreadsheetModel();
      const sheet = model.sheets[0];
      sheet.setCell(0, 0, { rawValue: 1, computedValue: 1 });
      sheet.setCell(0, 1, { rawValue: 3, computedValue: 3 });
      sheet.setCell(0, 2, { rawValue: 2, computedValue: 2 });

      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 2 };
      const cmd = new SortRangeCommand('Sheet1', range, 0, false);
      cmd.execute(model);
      expect(sheet.getCell(0, 0).rawValue).toBe(3);
      expect(sheet.getCell(0, 1).rawValue).toBe(2);
      expect(sheet.getCell(0, 2).rawValue).toBe(1);
    });

    it('sorts strings alphabetically', () => {
      const model = new SpreadsheetModel();
      const sheet = model.sheets[0];
      sheet.setCell(0, 0, { rawValue: 'banana', computedValue: 'banana' });
      sheet.setCell(0, 1, { rawValue: 'apple', computedValue: 'apple' });
      sheet.setCell(0, 2, { rawValue: 'cherry', computedValue: 'cherry' });

      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 2 };
      const cmd = new SortRangeCommand('Sheet1', range, 0, true);
      cmd.execute(model);
      expect(sheet.getCell(0, 0).rawValue).toBe('apple');
      expect(sheet.getCell(0, 2).rawValue).toBe('cherry');
    });

    it('places nulls last when sorting ascending', () => {
      const model = new SpreadsheetModel();
      const sheet = model.sheets[0];
      sheet.setCell(0, 0, { rawValue: 2, computedValue: 2 });
      sheet.setCell(0, 1, { rawValue: null, computedValue: null });
      sheet.setCell(0, 2, { rawValue: 1, computedValue: 1 });

      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 2 };
      const cmd = new SortRangeCommand('Sheet1', range, 0, true);
      cmd.execute(model);
      expect(sheet.getCell(0, 0).rawValue).toBe(1);
      expect(sheet.getCell(0, 1).rawValue).toBe(2);
    });
  });

  describe('SortRangeMultiCommand', () => {
    it('sorts by multiple keys', () => {
      // Col A: category, Col B: value
      // Data: (b,2), (a,3), (a,1)
      // Sort by col A asc, then col B asc → (a,1), (a,3), (b,2)
      const model = new SpreadsheetModel();
      const sheet = model.sheets[0];
      sheet.setCell(0, 0, { rawValue: 'b', computedValue: 'b' });
      sheet.setCell(1, 0, { rawValue: 2, computedValue: 2 });
      sheet.setCell(0, 1, { rawValue: 'a', computedValue: 'a' });
      sheet.setCell(1, 1, { rawValue: 3, computedValue: 3 });
      sheet.setCell(0, 2, { rawValue: 'a', computedValue: 'a' });
      sheet.setCell(1, 2, { rawValue: 1, computedValue: 1 });

      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 1, endRow: 2 };
      const cmd = new SortRangeMultiCommand('Sheet1', range, [
        { column: 0, ascending: true },
        { column: 1, ascending: true },
      ]);

      cmd.execute(model);
      expect(sheet.getCell(0, 0).rawValue).toBe('a');
      expect(sheet.getCell(1, 0).rawValue).toBe(1);
      expect(sheet.getCell(0, 1).rawValue).toBe('a');
      expect(sheet.getCell(1, 1).rawValue).toBe(3);
      expect(sheet.getCell(0, 2).rawValue).toBe('b');
    });

    it('undoes multi-key sort', () => {
      const model = new SpreadsheetModel();
      const sheet = model.sheets[0];
      sheet.setCell(0, 0, { rawValue: 3, computedValue: 3 });
      sheet.setCell(0, 1, { rawValue: 1, computedValue: 1 });

      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 1 };
      const cmd = new SortRangeMultiCommand('Sheet1', range, [{ column: 0, ascending: true }]);
      cmd.execute(model);
      cmd.undo(model);

      expect(sheet.getCell(0, 0).rawValue).toBe(3);
      expect(sheet.getCell(0, 1).rawValue).toBe(1);
    });
  });

  describe('PasteSpecialCommand', () => {
    it('pastes values only', () => {
      const model = new SpreadsheetModel();
      const cmd = new PasteSpecialCommand(
        'Sheet1',
        0,
        0,
        [
          { col: 0, row: 0, value: '42', styleId: null },
          { col: 1, row: 0, value: 'hello', styleId: null },
        ],
        'valuesOnly',
      );

      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBe('42');
      expect(model.sheets[0].getCell(1, 0).rawValue).toBe('hello');
    });

    it('pastes formula value', () => {
      const model = new SpreadsheetModel();
      const cmd = new PasteSpecialCommand(
        'Sheet1',
        0,
        0,
        [{ col: 0, row: 0, value: '=A2+B2', styleId: null }],
        'valuesOnly',
      );

      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).formula).toBe('A2+B2');
    });

    it('pastes formats only', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'keep', computedValue: 'keep' });
      const cmd = new PasteSpecialCommand(
        'Sheet1',
        0,
        0,
        [{ col: 0, row: 0, value: 'ignored', styleId: 'bold' }],
        'formatsOnly',
      );

      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBe('keep');
      expect(model.sheets[0].getCell(0, 0).styleId).toBe('bold');
    });

    it('pastes all (value + style)', () => {
      const model = new SpreadsheetModel();
      const cmd = new PasteSpecialCommand(
        'Sheet1',
        0,
        0,
        [{ col: 0, row: 0, value: 'data', styleId: 'italic' }],
        'all',
      );

      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBe('data');
      expect(model.sheets[0].getCell(0, 0).styleId).toBe('italic');
    });

    it('undoes paste', () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'original', computedValue: 'original' });
      const cmd = new PasteSpecialCommand(
        'Sheet1',
        0,
        0,
        [{ col: 0, row: 0, value: 'pasted', styleId: null }],
        'valuesOnly',
      );

      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBe('pasted');

      cmd.undo(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBe('original');
    });

    it('pastes at offset position', () => {
      const model = new SpreadsheetModel();
      const cmd = new PasteSpecialCommand(
        'Sheet1',
        3,
        2,
        [{ col: 0, row: 0, value: 'val', styleId: null }],
        'valuesOnly',
      );

      cmd.execute(model);
      expect(model.sheets[0].getCell(3, 2).rawValue).toBe('val');
    });

    it('treats empty string as null', () => {
      const model = new SpreadsheetModel();
      const cmd = new PasteSpecialCommand(
        'Sheet1',
        0,
        0,
        [{ col: 0, row: 0, value: '', styleId: null }],
        'valuesOnly',
      );

      cmd.execute(model);
      expect(model.sheets[0].getCell(0, 0).rawValue).toBeNull();
    });
  });
});
