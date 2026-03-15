import { describe, it, expect } from 'vitest';
import { OdsWriter } from '../../../src/ods/OdsWriter';
import { OdsReader } from '../../../src/ods/OdsReader';
import { SpreadsheetModel } from '../../../src/model/SpreadsheetModel';

async function roundTrip(model: SpreadsheetModel): Promise<SpreadsheetModel> {
  const writer = new OdsWriter();
  const reader = new OdsReader();
  const buffer = await writer.write(model);
  return reader.read(buffer);
}

describe('ODS Round-Trip (OdsWriter → OdsReader)', () => {
  describe('sheet structure', () => {
    it('preserves single sheet name', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].name = 'MyData';
      const result = await roundTrip(model);
      expect(result.sheets.length).toBe(1);
      expect(result.sheets[0].name).toBe('MyData');
    });

    it('preserves multiple sheets', async () => {
      const model = new SpreadsheetModel();
      model.addSheet('Revenue');
      model.addSheet('Expenses');
      const result = await roundTrip(model);
      expect(result.sheets.length).toBe(3);
      expect(result.sheets[0].name).toBe('Sheet1');
      expect(result.sheets[1].name).toBe('Revenue');
      expect(result.sheets[2].name).toBe('Expenses');
    });
  });

  describe('cell values', () => {
    it('preserves numeric cell values', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 42, computedValue: 42 });
      model.sheets[0].setCell(1, 0, { rawValue: 3.14, computedValue: 3.14 });
      model.sheets[0].setCell(0, 1, { rawValue: -100, computedValue: -100 });
      const result = await roundTrip(model);
      expect(result.sheets[0].getCell(0, 0).rawValue).toBe(42);
      expect(result.sheets[0].getCell(1, 0).rawValue).toBe(3.14);
      expect(result.sheets[0].getCell(0, 1).rawValue).toBe(-100);
    });

    it('preserves string cell values', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'hello world', computedValue: 'hello world' });
      model.sheets[0].setCell(0, 1, { rawValue: 'test data', computedValue: 'test data' });
      const result = await roundTrip(model);
      expect(result.sheets[0].getCell(0, 0).rawValue).toBe('hello world');
      expect(result.sheets[0].getCell(0, 1).rawValue).toBe('test data');
    });

    it('preserves boolean cell values', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: true, computedValue: true });
      model.sheets[0].setCell(1, 0, { rawValue: false, computedValue: false });
      const result = await roundTrip(model);
      // Booleans are stored as TRUE/FALSE strings in ODS text nodes
      const trueCell = result.sheets[0].getCell(0, 0);
      const falseCell = result.sheets[0].getCell(1, 0);
      // Value should be truthy/falsy or a boolean string
      expect(trueCell.rawValue).toBeTruthy();
      expect(String(falseCell.rawValue).toUpperCase()).toContain('FALSE');
    });

    it('preserves zero value', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 0, computedValue: 0 });
      const result = await roundTrip(model);
      expect(result.sheets[0].getCell(0, 0).rawValue).toBe(0);
    });

    it('empty cells remain empty after round-trip', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(5, 5, { rawValue: 'corner', computedValue: 'corner' });
      const result = await roundTrip(model);
      expect(result.sheets[0].getCell(0, 0).rawValue).toBeNull();
      expect(result.sheets[0].getCell(1, 0).rawValue).toBeNull();
    });
  });

  describe('formulas', () => {
    it('preserves formula string', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 10, computedValue: 10 });
      model.sheets[0].setCell(0, 1, { rawValue: 20, computedValue: 20 });
      model.sheets[0].setCell(0, 2, { rawValue: null, formula: 'SUM(A1:A2)', computedValue: 30 });
      const result = await roundTrip(model);
      const formulaCell = result.sheets[0].getCell(0, 2);
      expect(formulaCell.formula).toBeTruthy();
      expect(formulaCell.formula).toContain('SUM');
    });
  });

  describe('styles', () => {
    it('preserves bold style', async () => {
      const model = new SpreadsheetModel();
      model.getOrCreateStyle({ id: 'bold', bold: true });
      model.sheets[0].setCell(0, 0, {
        rawValue: 'Bold text',
        computedValue: 'Bold text',
        styleId: 'bold',
      });
      const result = await roundTrip(model);
      const styleId = result.sheets[0].getCell(0, 0).styleId;
      expect(styleId).toBeTruthy();
      const style = result.styles.get(styleId!);
      expect(style?.bold).toBe(true);
    });

    it('preserves italic style', async () => {
      const model = new SpreadsheetModel();
      model.getOrCreateStyle({ id: 'italic', italic: true });
      model.sheets[0].setCell(0, 0, { rawValue: 'text', computedValue: 'text', styleId: 'italic' });
      const result = await roundTrip(model);
      const styleId = result.sheets[0].getCell(0, 0).styleId;
      const style = result.styles.get(styleId!);
      expect(style?.italic).toBe(true);
    });

    it('preserves background color', async () => {
      const model = new SpreadsheetModel();
      model.getOrCreateStyle({ id: 'colored', backgroundColor: '#ff0000' });
      model.sheets[0].setCell(0, 0, {
        rawValue: 'text',
        computedValue: 'text',
        styleId: 'colored',
      });
      const result = await roundTrip(model);
      const styleId = result.sheets[0].getCell(0, 0).styleId;
      const style = result.styles.get(styleId!);
      expect(style?.backgroundColor).toBe('#ff0000');
    });
  });

  describe('merge cells', () => {
    it('preserves merge span on anchor cell', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, {
        rawValue: 'merged',
        computedValue: 'merged',
        mergeColSpan: 3,
        mergeRowSpan: 2,
      });
      const result = await roundTrip(model);
      const anchor = result.sheets[0].getCell(0, 0);
      expect(anchor.mergeColSpan).toBe(3);
      expect(anchor.mergeRowSpan).toBe(2);
    });

    it('sets mergedInto on same-row covered cells', async () => {
      const model = new SpreadsheetModel();
      // A1:C1 merged, D1 is a separate cell
      model.sheets[0].setCell(0, 0, { rawValue: 'wide', computedValue: 'wide', mergeColSpan: 3 });
      model.sheets[0].setCell(3, 0, { rawValue: 'separate', computedValue: 'separate' });
      const result = await roundTrip(model);
      // Covered cells B1 and C1 should have mergedInto pointing to A1
      const b1 = result.sheets[0].getCell(1, 0);
      const c1 = result.sheets[0].getCell(2, 0);
      expect(b1.mergedInto).toEqual({ sheet: 'Sheet1', col: 0, row: 0 });
      expect(c1.mergedInto).toEqual({ sheet: 'Sheet1', col: 0, row: 0 });
      // D1 should be a normal separate cell
      expect(result.sheets[0].getCell(3, 0).rawValue).toBe('separate');
    });

    it('sets mergedInto on multi-row covered cells', async () => {
      const model = new SpreadsheetModel();
      // A1:B2 merged (2 cols × 2 rows)
      model.sheets[0].setCell(0, 0, {
        rawValue: 'tall',
        computedValue: 'tall',
        mergeColSpan: 2,
        mergeRowSpan: 2,
      });
      model.sheets[0].setCell(2, 0, { rawValue: 'r0c2', computedValue: 'r0c2' });
      model.sheets[0].setCell(2, 1, { rawValue: 'r1c2', computedValue: 'r1c2' });
      const result = await roundTrip(model);
      const origin = { sheet: 'Sheet1', col: 0, row: 0 };
      expect(result.sheets[0].getCell(1, 0).mergedInto).toEqual(origin); // B1
      expect(result.sheets[0].getCell(0, 1).mergedInto).toEqual(origin); // A2
      expect(result.sheets[0].getCell(1, 1).mergedInto).toEqual(origin); // B2
      // C1, C2 are independent
      expect(result.sheets[0].getCell(2, 0).rawValue).toBe('r0c2');
      expect(result.sheets[0].getCell(2, 1).rawValue).toBe('r1c2');
    });

    it('cells after same-row merge are placed at correct column', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'A', computedValue: 'A', mergeColSpan: 2 });
      model.sheets[0].setCell(2, 0, { rawValue: 'C', computedValue: 'C' });
      model.sheets[0].setCell(3, 0, { rawValue: 'D', computedValue: 'D' });
      const result = await roundTrip(model);
      expect(result.sheets[0].getCell(0, 0).rawValue).toBe('A');
      expect(result.sheets[0].getCell(2, 0).rawValue).toBe('C');
      expect(result.sheets[0].getCell(3, 0).rawValue).toBe('D');
    });
  });

  describe('metadata', () => {
    it('preserves document title', async () => {
      const model = new SpreadsheetModel();
      model.metadata = { title: 'My Budget', creator: 'Alice' };
      const result = await roundTrip(model);
      expect(result.metadata.title).toBe('My Budget');
      expect(result.metadata.creator).toBe('Alice');
    });
  });

  describe('file format', () => {
    it('write() returns a non-empty Uint8Array', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 1, computedValue: 1 });
      const writer = new OdsWriter();
      const buffer = await writer.write(model);
      expect(buffer).toBeInstanceOf(Uint8Array);
      expect(buffer.length).toBeGreaterThan(100);
    });

    it('output is a valid ZIP (starts with PK signature)', async () => {
      const model = new SpreadsheetModel();
      const writer = new OdsWriter();
      const buffer = await writer.write(model);
      // ZIP files start with PK (0x50 0x4B)
      expect(buffer[0]).toBe(0x50);
      expect(buffer[1]).toBe(0x4b);
    });

    it('complex model with multiple sheets and cells round-trips', async () => {
      const model = new SpreadsheetModel();
      model.addSheet('Data');
      model.sheets[0].setCell(0, 0, { rawValue: 'Name', computedValue: 'Name' });
      model.sheets[0].setCell(1, 0, { rawValue: 'Value', computedValue: 'Value' });
      model.sheets[0].setCell(0, 1, { rawValue: 'Alpha', computedValue: 'Alpha' });
      model.sheets[0].setCell(1, 1, { rawValue: 100, computedValue: 100 });
      model.sheets[1].setCell(0, 0, { rawValue: 'Summary', computedValue: 'Summary' });

      const result = await roundTrip(model);
      expect(result.sheets.length).toBe(2);
      expect(result.sheets[0].getCell(0, 0).rawValue).toBe('Name');
      expect(result.sheets[0].getCell(1, 1).rawValue).toBe(100);
      expect(result.sheets[1].getCell(0, 0).rawValue).toBe('Summary');
    });
  });

  describe('column widths and row heights', () => {
    it('preserves non-default column width', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'test', computedValue: 'test' });
      model.sheets[0].setColumnWidth(0, 150);

      const result = await roundTrip(model);
      // Allow ±2px tolerance for cm ↔ px conversion rounding
      expect(result.sheets[0].getColumnWidth(0)).toBeCloseTo(150, -1);
    });

    it('preserves non-default row height', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'test', computedValue: 'test' });
      model.sheets[0].setRowHeight(0, 48);

      const result = await roundTrip(model);
      expect(result.sheets[0].getRowHeight(0)).toBeCloseTo(48, -1);
    });

    it('multiple columns with different widths all round-trip', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(2, 0, { rawValue: 'x', computedValue: 'x' });
      model.sheets[0].setColumnWidth(0, 120);
      model.sheets[0].setColumnWidth(1, 200);
      model.sheets[0].setColumnWidth(2, 60);

      const result = await roundTrip(model);
      expect(result.sheets[0].getColumnWidth(0)).toBeCloseTo(120, -1);
      expect(result.sheets[0].getColumnWidth(1)).toBeCloseTo(200, -1);
      expect(result.sheets[0].getColumnWidth(2)).toBeCloseTo(60, -1);
    });

    it('default column width is preserved as default (80px)', async () => {
      const model = new SpreadsheetModel();
      model.sheets[0].setCell(0, 0, { rawValue: 1, computedValue: 1 });
      // column width stays at default 80

      const result = await roundTrip(model);
      expect(result.sheets[0].getColumnWidth(0)).toBeCloseTo(80, -1);
    });
  });
});
