/**
 * Integration tests: full pipeline from raw user input → model mutation →
 * formula recalculation → serialization.
 *
 * These tests simulate what odsEditorProvider does when it receives 'editCell'
 * messages from the webview, verifying that the entire chain works correctly.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { SpreadsheetModel } from '../../../src/model/SpreadsheetModel';
import { FormulaEngine } from '../../../src/formula/FormulaEngine';
import { SetCellValueCommand } from '../../../src/model/commands';

// ── helpers ──────────────────────────────────────────────────────────────────

/** Simulates the value-parsing logic from odsEditorProvider.handleWebviewMessage */
function parseUserInput(
  raw: string,
): { value: string | number | boolean | null; formula: string | null } {
  if (raw.startsWith('=')) return { value: null, formula: raw.substring(1) };
  if (raw === '') return { value: null, formula: null };
  const num = Number(raw);
  if (!isNaN(num) && isFinite(num)) return { value: num, formula: null };
  if (raw.toUpperCase() === 'TRUE') return { value: true, formula: null };
  if (raw.toUpperCase() === 'FALSE') return { value: false, formula: null };
  return { value: raw, formula: null };
}

function makeModelAndEngine(): { model: SpreadsheetModel; engine: FormulaEngine } {
  const model = new SpreadsheetModel();
  const engine = new FormulaEngine();
  return { model, engine };
}

/** Apply an edit as odsEditorProvider would */
function editCell(
  model: SpreadsheetModel,
  engine: FormulaEngine,
  sheetName: string,
  col: number,
  row: number,
  rawInput: string,
): void {
  const sheet = model.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

  const { value, formula } = parseUserInput(rawInput);
  const cmd = new SetCellValueCommand(sheetName, col, row, value, formula);
  cmd.execute(model);
  engine.recalcAll(model);
}

// ── tests ─────────────────────────────────────────────────────────────────────

describe('Cell edit → formula recalculation pipeline', () => {
  let model: SpreadsheetModel;
  let engine: FormulaEngine;

  beforeEach(() => {
    ({ model, engine } = makeModelAndEngine());
  });

  const sheet = () => model.sheets[0].name;

  it('stores a plain number', () => {
    editCell(model, engine, sheet(), 0, 0, '42');
    const cell = model.sheets[0].getCell(0, 0);
    expect(cell.rawValue).toBe(42);
    expect(cell.formula).toBeNull();
  });

  it('stores a plain string', () => {
    editCell(model, engine, sheet(), 0, 0, 'Hello');
    expect(model.sheets[0].getCell(0, 0).rawValue).toBe('Hello');
  });

  it('stores boolean TRUE', () => {
    editCell(model, engine, sheet(), 0, 0, 'TRUE');
    expect(model.sheets[0].getCell(0, 0).rawValue).toBe(true);
  });

  it('clears a cell on empty input', () => {
    editCell(model, engine, sheet(), 0, 0, '99');
    editCell(model, engine, sheet(), 0, 0, '');
    expect(model.sheets[0].getCell(0, 0).rawValue).toBeNull();
  });

  it('evaluates a SUM formula after editing source cells', () => {
    editCell(model, engine, sheet(), 0, 0, '10');
    editCell(model, engine, sheet(), 0, 1, '20');
    editCell(model, engine, sheet(), 0, 2, '=SUM(A1:A2)');
    expect(model.sheets[0].getCell(0, 2).computedValue).toBe(30);
  });

  it('recalculates downstream formulas when a source cell changes', () => {
    editCell(model, engine, sheet(), 0, 0, '5');
    editCell(model, engine, sheet(), 0, 1, '=A1*2');
    editCell(model, engine, sheet(), 0, 2, '=A2+1');

    expect(model.sheets[0].getCell(0, 1).computedValue).toBe(10);
    expect(model.sheets[0].getCell(0, 2).computedValue).toBe(11);

    // Now change source
    editCell(model, engine, sheet(), 0, 0, '10');
    expect(model.sheets[0].getCell(0, 1).computedValue).toBe(20);
    expect(model.sheets[0].getCell(0, 2).computedValue).toBe(21);
  });

  it('handles IF formula with boolean result', () => {
    editCell(model, engine, sheet(), 0, 0, '100');
    editCell(model, engine, sheet(), 0, 1, '=IF(A1>50,"Big","Small")');
    expect(model.sheets[0].getCell(0, 1).computedValue).toBe('Big');
  });

  it('detects circular references', () => {
    editCell(model, engine, sheet(), 0, 0, '=B1');
    editCell(model, engine, sheet(), 1, 0, '=A1');
    // At least one cell should have #CIRC! error
    const a1 = model.sheets[0].getCell(0, 0).computedValue;
    const b1 = model.sheets[0].getCell(1, 0).computedValue;
    expect(String(a1) === '#CIRC!' || String(b1) === '#CIRC!').toBe(true);
  });

  it('handles cross-sheet formula references', () => {
    model.addSheet('Data');
    const dataSheet = model.getSheetByName('Data');
    if (!dataSheet) throw new Error('Data sheet not found');
    dataSheet.setCell(0, 0, { rawValue: 99, computedValue: 99 });

    editCell(model, engine, sheet(), 0, 0, '=Data.A1');
    expect(model.sheets[0].getCell(0, 0).computedValue).toBe(99);
  });
});

describe('Undo/Redo pipeline', () => {
  let model: SpreadsheetModel;
  let engine: FormulaEngine;

  beforeEach(() => {
    ({ model, engine } = makeModelAndEngine());
  });

  const sheet = () => model.sheets[0].name;

  it('undo restores previous value', () => {
    const cmd = new SetCellValueCommand(sheet(), 0, 0, 'first', null);
    cmd.execute(model);

    const cmd2 = new SetCellValueCommand(sheet(), 0, 0, 'second', null);
    cmd2.execute(model);
    expect(model.sheets[0].getCell(0, 0).rawValue).toBe('second');

    cmd2.undo(model);
    expect(model.sheets[0].getCell(0, 0).rawValue).toBe('first');
  });

  it('redo reapplies the edit', () => {
    const cmd = new SetCellValueCommand(sheet(), 0, 0, 42, null);
    cmd.execute(model);
    cmd.undo(model);
    expect(model.sheets[0].getCell(0, 0).rawValue).toBeNull();

    cmd.execute(model);
    expect(model.sheets[0].getCell(0, 0).rawValue).toBe(42);
  });
});

describe('Multi-sheet model', () => {
  it('adding a sheet increments sheets count', () => {
    const model = new SpreadsheetModel();
    expect(model.sheets.length).toBe(1);
    model.addSheet('Sheet2');
    expect(model.sheets.length).toBe(2);
    expect(model.sheets[1].name).toBe('Sheet2');
  });

  it('removing a sheet decrements count', () => {
    const model = new SpreadsheetModel();
    model.addSheet('Sheet2');
    model.removeSheet(1);
    expect(model.sheets.length).toBe(1);
  });

  it('renaming a sheet updates the name', () => {
    const model = new SpreadsheetModel();
    model.renameSheet(0, 'Budget');
    expect(model.sheets[0].name).toBe('Budget');
  });

  it('reordering sheets changes their positions', () => {
    const model = new SpreadsheetModel();
    model.addSheet('B');
    model.addSheet('C');
    model.reorderSheet(2, 0); // move C to front
    expect(model.sheets[0].name).toBe('C');
  });
});

describe('Named ranges', () => {
  it('can define and delete a named range', () => {
    const model = new SpreadsheetModel();
    model.defineNamedRange('SALES', 'Sheet1', { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 9 });
    expect(model.namedRanges.has('SALES')).toBe(true);
    model.deleteNamedRange('SALES');
    expect(model.namedRanges.has('SALES')).toBe(false);
  });

  it('formula engine resolves named ranges', () => {
    const model = new SpreadsheetModel();
    const engine = new FormulaEngine();
    model.sheets[0].setCell(0, 0, { rawValue: 10, computedValue: 10 });
    model.sheets[0].setCell(0, 1, { rawValue: 20, computedValue: 20 });
    model.defineNamedRange('MYRANGE', 'Sheet1', { sheet: 'Sheet1', startCol: 0, startRow: 0, endCol: 0, endRow: 1 });
    model.sheets[0].setCell(0, 2, { rawValue: null, formula: 'SUM(MYRANGE)', computedValue: null });
    engine.recalcAll(model);
    expect(model.sheets[0].getCell(0, 2).computedValue).toBe(30);
  });
});
