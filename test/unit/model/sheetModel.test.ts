import { describe, it, expect } from 'vitest';
import { SheetModel } from '../../../src/model/SheetModel';

describe('SheetModel', () => {
  it('creates a sheet with default dimensions', () => {
    const sheet = new SheetModel('Test');
    expect(sheet.name).toBe('Test');
    expect(sheet.columnCount).toBe(1024);
    expect(sheet.rowCount).toBe(65536);
    expect(sheet.cellCount).toBe(0);
  });

  it('gets empty cell for unpopulated position', () => {
    const sheet = new SheetModel('Test');
    const cell = sheet.getCell(0, 0);
    expect(cell.rawValue).toBeNull();
    expect(cell.formula).toBeNull();
  });

  it('sets and gets cell data', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 42, computedValue: 42 });
    const cell = sheet.getCell(0, 0);
    expect(cell.rawValue).toBe(42);
    expect(cell.computedValue).toBe(42);
    expect(sheet.cellCount).toBe(1);
  });

  it('removes cell when set to empty', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 42, computedValue: 42 });
    expect(sheet.cellCount).toBe(1);
    sheet.setCell(0, 0, {
      rawValue: null,
      formula: null,
      computedValue: null,
      styleId: null,
      mergeColSpan: 1,
      mergeRowSpan: 1,
      mergedInto: null,
    });
    expect(sheet.cellCount).toBe(0);
  });

  it('hasCell returns false for unpopulated position', () => {
    const sheet = new SheetModel('Test');
    expect(sheet.hasCell(0, 0)).toBe(false);
    sheet.setCell(0, 0, { rawValue: 1, computedValue: 1 });
    expect(sheet.hasCell(0, 0)).toBe(true);
  });

  it('deleteCell removes cell', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 1, computedValue: 1 });
    sheet.deleteCell(0, 0);
    expect(sheet.hasCell(0, 0)).toBe(false);
    expect(sheet.cellCount).toBe(0);
  });

  it('partial setCell merges with existing data', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 'hello', computedValue: 'hello', styleId: 'bold' });
    // Only update styleId
    sheet.setCell(0, 0, { styleId: 'italic' });
    expect(sheet.getCell(0, 0).rawValue).toBe('hello');
    expect(sheet.getCell(0, 0).styleId).toBe('italic');
  });

  it('returns used range', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(2, 3, { rawValue: 'a', computedValue: 'a' });
    sheet.setCell(5, 1, { rawValue: 'b', computedValue: 'b' });
    const range = sheet.getUsedRange();
    expect(range).not.toBeNull();
    expect(range!.startCol).toBe(2);
    expect(range!.endCol).toBe(5);
    expect(range!.startRow).toBe(1);
    expect(range!.endRow).toBe(3);
  });

  it('returns null for empty sheet used range', () => {
    const sheet = new SheetModel('Test');
    expect(sheet.getUsedRange()).toBeNull();
  });

  it('inserts rows and shifts cells', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 'a', computedValue: 'a' });
    sheet.setCell(0, 1, { rawValue: 'b', computedValue: 'b' });
    sheet.insertRows(1, 2);
    expect(sheet.getCell(0, 0).rawValue).toBe('a');
    expect(sheet.getCell(0, 1).rawValue).toBeNull();
    expect(sheet.getCell(0, 3).rawValue).toBe('b');
  });

  it('insertRows increases rowCount', () => {
    const sheet = new SheetModel('Test');
    const orig = sheet.rowCount;
    sheet.insertRows(0, 5);
    expect(sheet.rowCount).toBe(orig + 5);
  });

  it('deletes rows and shifts cells', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 'a', computedValue: 'a' });
    sheet.setCell(0, 2, { rawValue: 'c', computedValue: 'c' });
    sheet.deleteRows(1, 1);
    expect(sheet.getCell(0, 0).rawValue).toBe('a');
    expect(sheet.getCell(0, 1).rawValue).toBe('c');
  });

  it('deleteRows removes cells in deleted range', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 1, { rawValue: 'deleted', computedValue: 'deleted' });
    sheet.deleteRows(1, 1);
    expect(sheet.getCell(0, 1).rawValue).toBeNull();
  });

  it('deleteRows decreases rowCount', () => {
    const sheet = new SheetModel('Test');
    const orig = sheet.rowCount;
    sheet.deleteRows(0, 3);
    expect(sheet.rowCount).toBe(orig - 3);
  });

  it('inserts columns and shifts cells', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 'a', computedValue: 'a' });
    sheet.setCell(1, 0, { rawValue: 'b', computedValue: 'b' });
    sheet.insertColumns(1, 1);
    expect(sheet.getCell(0, 0).rawValue).toBe('a');
    expect(sheet.getCell(1, 0).rawValue).toBeNull();
    expect(sheet.getCell(2, 0).rawValue).toBe('b');
  });

  it('insertColumns increases columnCount', () => {
    const sheet = new SheetModel('Test');
    const orig = sheet.columnCount;
    sheet.insertColumns(0, 3);
    expect(sheet.columnCount).toBe(orig + 3);
  });

  it('deletes columns and shifts cells', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 'a', computedValue: 'a' });
    sheet.setCell(2, 0, { rawValue: 'c', computedValue: 'c' });
    sheet.deleteColumns(1, 1);
    expect(sheet.getCell(0, 0).rawValue).toBe('a');
    expect(sheet.getCell(1, 0).rawValue).toBe('c');
  });

  it('deleteColumns decreases columnCount', () => {
    const sheet = new SheetModel('Test');
    const orig = sheet.columnCount;
    sheet.deleteColumns(0, 2);
    expect(sheet.columnCount).toBe(orig - 2);
  });

  it('sets column width with minimum enforcement', () => {
    const sheet = new SheetModel('Test');
    sheet.setColumnWidth(0, 5);
    expect(sheet.getColumnWidth(0)).toBe(20); // minimum is 20
    sheet.setColumnWidth(0, 200);
    expect(sheet.getColumnWidth(0)).toBe(200);
  });

  it('sets row height with minimum enforcement', () => {
    const sheet = new SheetModel('Test');
    sheet.setRowHeight(0, 2);
    expect(sheet.getRowHeight(0)).toBe(10); // minimum is 10
    sheet.setRowHeight(0, 50);
    expect(sheet.getRowHeight(0)).toBe(50);
  });

  it('setColumnWidth does nothing for out-of-range column', () => {
    const sheet = new SheetModel('Test', 5, 10);
    expect(() => sheet.setColumnWidth(100, 80)).not.toThrow();
  });

  it('setRowHeight does nothing for out-of-range row', () => {
    const sheet = new SheetModel('Test', 5, 10);
    expect(() => sheet.setRowHeight(100, 24)).not.toThrow();
  });

  it('gets cell range', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 1, computedValue: 1 });
    sheet.setCell(1, 0, { rawValue: 2, computedValue: 2 });
    sheet.setCell(0, 1, { rawValue: 3, computedValue: 3 });
    sheet.setCell(1, 1, { rawValue: 4, computedValue: 4 });
    const range = sheet.getCellRange({
      sheet: 'Test',
      startCol: 0,
      startRow: 0,
      endCol: 1,
      endRow: 1,
    });
    expect(range.length).toBe(2);
    expect(range[0].length).toBe(2);
    expect(range[0][0].rawValue).toBe(1);
    expect(range[1][1].rawValue).toBe(4);
  });

  it('getAllCells returns all non-empty cells', () => {
    const sheet = new SheetModel('Test');
    sheet.setCell(0, 0, { rawValue: 1, computedValue: 1 });
    sheet.setCell(5, 5, { rawValue: 2, computedValue: 2 });
    const all = sheet.getAllCells();
    expect(all.length).toBe(2);
    const keys = all.map((c) => `${c.col},${c.row}`).sort();
    expect(keys).toContain('0,0');
    expect(keys).toContain('5,5');
  });

  it('insertRows inserts default row heights', () => {
    const sheet = new SheetModel('Test');
    sheet.setRowHeight(0, 50);
    sheet.insertRows(0, 1);
    expect(sheet.getRowHeight(0)).toBe(24); // default inserted
    expect(sheet.getRowHeight(1)).toBe(50); // old row 0 shifted
  });

  it('deleteRows removes row heights', () => {
    const sheet = new SheetModel('Test');
    sheet.setRowHeight(1, 99);
    sheet.deleteRows(1, 1);
    // row 1 deleted; former row 2 (default) is now row 1
    expect(sheet.getRowHeight(1)).toBe(24);
  });

  describe('freeze panes defaults', () => {
    it('starts with no frozen rows or cols', () => {
      const sheet = new SheetModel('Test');
      expect(sheet.frozenRows).toBe(0);
      expect(sheet.frozenCols).toBe(0);
    });
  });

  describe('filters', () => {
    it('starts with no filters', () => {
      const sheet = new SheetModel('Test');
      expect(sheet.filters.size).toBe(0);
    });
  });

  describe('hiddenRows', () => {
    it('starts with no hidden rows', () => {
      const sheet = new SheetModel('Test');
      expect(sheet.hiddenRows.size).toBe(0);
    });
  });
});
