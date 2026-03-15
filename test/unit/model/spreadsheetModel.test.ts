import { describe, it, expect } from 'vitest';
import { SpreadsheetModel } from '../../../src/model/SpreadsheetModel';

describe('SpreadsheetModel', () => {
  it('creates with one default sheet', () => {
    const model = new SpreadsheetModel();
    expect(model.sheets.length).toBe(1);
    expect(model.sheets[0].name).toBe('Sheet1');
  });

  it('activeSheet returns the active sheet', () => {
    const model = new SpreadsheetModel();
    expect(model.activeSheet).toBe(model.sheets[0]);
    model.addSheet('Sheet2');
    model.activeSheetIndex = 1;
    expect(model.activeSheet.name).toBe('Sheet2');
  });

  it('adds sheets', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    expect(model.sheets.length).toBe(2);
    expect(model.sheets[1].name).toBe('Sheet2');
  });

  it('addSheet returns the new sheet', () => {
    const model = new SpreadsheetModel();
    const sheet = model.addSheet('Data');
    expect(sheet).toBe(model.sheets[1]);
    expect(sheet.name).toBe('Data');
  });

  it('removes sheets but not the last one', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.removeSheet(1);
    expect(model.sheets.length).toBe(1);
    // Cannot remove last sheet
    model.removeSheet(0);
    expect(model.sheets.length).toBe(1);
  });

  it('removeSheet adjusts activeSheetIndex when needed', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.addSheet('Sheet3');
    model.activeSheetIndex = 2;
    model.removeSheet(2);
    // Active index should clamp to last existing sheet
    expect(model.activeSheetIndex).toBe(1);
  });

  it('renames sheets', () => {
    const model = new SpreadsheetModel();
    model.renameSheet(0, 'Data');
    expect(model.sheets[0].name).toBe('Data');
  });

  it('renameSheet ignores out-of-range index', () => {
    const model = new SpreadsheetModel();
    expect(() => model.renameSheet(99, 'Bad')).not.toThrow();
    expect(model.sheets[0].name).toBe('Sheet1');
  });

  it('reorders sheets', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.addSheet('Sheet3');
    model.activeSheetIndex = 0;
    model.reorderSheet(0, 2);
    expect(model.sheets[0].name).toBe('Sheet2');
    expect(model.sheets[2].name).toBe('Sheet1');
    expect(model.activeSheetIndex).toBe(2);
  });

  it('reorderSheet adjusts activeSheetIndex when active moves right', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.addSheet('Sheet3');
    model.activeSheetIndex = 2; // Sheet3
    // Move Sheet1 (index 0) to index 2 — active (Sheet3) was at 2, it shifts left
    model.reorderSheet(0, 2);
    expect(model.activeSheetIndex).toBe(1);
  });

  it('reorderSheet adjusts activeSheetIndex when active moves left', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.addSheet('Sheet3');
    model.activeSheetIndex = 0; // Sheet1
    // Move Sheet3 (index 2) to index 0 — active (Sheet1) was at 0, it shifts right
    model.reorderSheet(2, 0);
    expect(model.activeSheetIndex).toBe(1);
  });

  it('reorderSheet does nothing for out-of-range indices', () => {
    const model = new SpreadsheetModel();
    expect(() => model.reorderSheet(-1, 0)).not.toThrow();
    expect(() => model.reorderSheet(0, 99)).not.toThrow();
    expect(model.sheets.length).toBe(1);
  });

  it('finds sheets by name', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Data');
    const sheet = model.getSheetByName('Data');
    expect(sheet).not.toBeUndefined();
    expect(sheet!.name).toBe('Data');
    expect(model.getSheetByName('NonExistent')).toBeUndefined();
  });

  it('manages styles', () => {
    const model = new SpreadsheetModel();
    const style = model.getOrCreateStyle({ id: 's1', bold: true });
    expect(style.id).toBe('s1');
    expect(style.bold).toBe(true);

    // Retrieve same style — not updated, existing returned
    const same = model.getOrCreateStyle({ id: 's1', italic: true });
    expect(same.bold).toBe(true);
    expect(same.italic).toBeUndefined();
  });

  it('resolveStyle returns undefined for unknown id', () => {
    const model = new SpreadsheetModel();
    expect(model.resolveStyle(null)).toBeUndefined();
    expect(model.resolveStyle('unknown')).toBeUndefined();
  });

  it('resolveStyle returns known style', () => {
    const model = new SpreadsheetModel();
    model.getOrCreateStyle({ id: 'bold', bold: true });
    expect(model.resolveStyle('bold')?.bold).toBe(true);
  });

  describe('named ranges', () => {
    it('defines and resolves a named range', () => {
      const model = new SpreadsheetModel();
      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 2, endRow: 5 };
      model.defineNamedRange('MyRange', 'Sheet1', range);

      const resolved = model.resolveNamedRange('MyRange');
      expect(resolved).not.toBeUndefined();
      expect(resolved!.name).toBe('MyRange');
      expect(resolved!.sheet).toBe('Sheet1');
      expect(resolved!.range.endCol).toBe(2);
    });

    it('resolves named range case-insensitively', () => {
      const model = new SpreadsheetModel();
      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 0 };
      model.defineNamedRange('Total', 'Sheet1', range);
      expect(model.resolveNamedRange('TOTAL')).not.toBeUndefined();
      expect(model.resolveNamedRange('total')).not.toBeUndefined();
    });

    it('deletes named range', () => {
      const model = new SpreadsheetModel();
      const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 0 };
      model.defineNamedRange('Temp', 'Sheet1', range);
      model.deleteNamedRange('Temp');
      expect(model.resolveNamedRange('Temp')).toBeUndefined();
    });

    it('returns undefined for unknown named range', () => {
      const model = new SpreadsheetModel();
      expect(model.resolveNamedRange('Unknown')).toBeUndefined();
    });

    it('overwrites existing named range', () => {
      const model = new SpreadsheetModel();
      const r1 = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 1, endRow: 1 };
      const r2 = { sheet: 'Sheet1', startCol: 5, startRow: 5, endCol: 6, endRow: 6 };
      model.defineNamedRange('MyRange', 'Sheet1', r1);
      model.defineNamedRange('MyRange', 'Sheet1', r2);
      const resolved = model.resolveNamedRange('MyRange');
      expect(resolved!.range.startCol).toBe(5);
    });
  });

  it('serializes the model', () => {
    const model = new SpreadsheetModel();
    model.sheets[0].setCell(0, 0, { rawValue: 42, computedValue: 42 });
    model.getOrCreateStyle({ id: 's1', bold: true });

    const serialized = model.serialize();
    expect(serialized.sheets.length).toBe(1);
    expect(serialized.sheets[0].cells.length).toBe(1);
    expect(serialized.sheets[0].cells[0].data.rawValue).toBe(42);
    expect(serialized.styles['s1'].bold).toBe(true);
  });

  it('serializes named ranges', () => {
    const model = new SpreadsheetModel();
    const range = { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 0 };
    model.defineNamedRange('Budget', 'Sheet1', range);

    const serialized = model.serialize();
    expect(serialized.namedRanges).toBeDefined();
    expect(serialized.namedRanges!.length).toBe(1);
    expect(serialized.namedRanges![0].name).toBe('Budget');
  });

  it('serializes multiple sheets', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.sheets[1].setCell(0, 0, { rawValue: 'hello', computedValue: 'hello' });

    const serialized = model.serialize();
    expect(serialized.sheets.length).toBe(2);
    expect(serialized.sheets[1].cells[0].data.rawValue).toBe('hello');
  });

  it('serializes activeSheetIndex', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.activeSheetIndex = 1;

    const serialized = model.serialize();
    expect(serialized.activeSheetIndex).toBe(1);
  });

  it('serializes frozen panes', () => {
    const model = new SpreadsheetModel();
    model.sheets[0].frozenRows = 2;
    model.sheets[0].frozenCols = 1;

    const serialized = model.serialize();
    expect(serialized.sheets[0].frozenRows).toBe(2);
    expect(serialized.sheets[0].frozenCols).toBe(1);
  });
});
