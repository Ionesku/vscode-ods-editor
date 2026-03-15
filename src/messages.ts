import { CellData, CellRange, CellStyle, DataValidation, FilterCriteria } from './model/types';
import { SerializedModel, SerializedSheet } from './model/SpreadsheetModel';

// ── Extension -> Webview ────────────────────────────────────────────────────

export type ExtToWebviewMessage =
  | { type: 'init'; data: SerializedModel }
  | {
      type: 'cellsUpdated';
      cells: Array<{ sheet: string; col: number; row: number; data: CellData }>;
    }
  | { type: 'styleUpdated'; styleId: string; style: CellStyle }
  | { type: 'sheetAdded'; sheet: SerializedSheet; index: number }
  | { type: 'sheetRemoved'; index: number }
  | { type: 'sheetRenamed'; index: number; name: string }
  | { type: 'modelSnapshot'; data: SerializedModel }
  | { type: 'columnValues'; column: number; values: string[] };

// ── Webview -> Extension ────────────────────────────────────────────────────

export type WebviewToExtMessage =
  | { type: 'ready' }
  | { type: 'editCell'; sheet: string; col: number; row: number; value: string }
  | { type: 'setStyle'; sheet: string; range: CellRange; style: Partial<CellStyle> }
  | { type: 'mergeCells'; sheet: string; range: CellRange }
  | { type: 'unmergeCells'; sheet: string; range: CellRange }
  | { type: 'resizeColumn'; sheet: string; col: number; width: number }
  | { type: 'resizeRow'; sheet: string; row: number; height: number }
  | { type: 'addSheet'; name: string }
  | { type: 'removeSheet'; index: number }
  | { type: 'renameSheet'; index: number; name: string }
  | { type: 'switchSheet'; index: number }
  | { type: 'sort'; sheet: string; range: CellRange; column: number; ascending: boolean }
  | {
      type: 'sortMulti';
      sheet: string;
      range: CellRange;
      keys: Array<{ column: number; ascending: boolean }>;
    }
  | { type: 'setFilter'; sheet: string; column: number; criteria: FilterCriteria | null }
  | { type: 'insertRows'; sheet: string; at: number; count: number }
  | { type: 'insertColumns'; sheet: string; at: number; count: number }
  | { type: 'deleteRows'; sheet: string; at: number; count: number }
  | { type: 'deleteColumns'; sheet: string; at: number; count: number }
  | { type: 'freezePanes'; sheet: string; frozenRows: number; frozenCols: number }
  | { type: 'autoFitColumn'; sheet: string; col: number }
  | { type: 'getColumnValues'; sheet: string; column: number }
  | { type: 'setConditionalFormat'; sheet: string; rule: ConditionalFormatRule }
  | { type: 'removeConditionalFormat'; sheet: string; ruleId: string }
  | {
      type: 'pasteSpecial';
      sheet: string;
      startCol: number;
      startRow: number;
      cells: PasteCell[];
      mode: 'valuesOnly' | 'formatsOnly' | 'all';
    }
  | { type: 'reorderSheet'; fromIndex: number; toIndex: number }
  | {
      type: 'defineNamedRange';
      name: string;
      sheet: string;
      range: import('./model/types').CellRange;
    }
  | { type: 'deleteNamedRange'; name: string }
  | { type: 'setDataValidation'; sheet: string; validation: DataValidation }
  | { type: 'removeDataValidation'; sheet: string; id: string }
  | { type: 'moveRange'; sheet: string; fromRange: CellRange; toCol: number; toRow: number }
  | { type: 'setComment'; sheet: string; col: number; row: number; comment: string };

export interface PasteCell {
  col: number; // offset from paste origin
  row: number;
  value: string; // raw value / formula
  styleId: string | null;
}

export interface ConditionalFormatRule {
  id: string;
  range: CellRange;
  condition:
    | 'greaterThan'
    | 'lessThan'
    | 'equals'
    | 'notEquals'
    | 'contains'
    | 'notContains'
    | 'between'
    | 'empty'
    | 'notEmpty'
    | 'colorScale'
    | 'dataBar';
  value1: string;
  value2?: string; // for 'between'
  style: { backgroundColor?: string; textColor?: string; bold?: boolean; italic?: boolean };
  /** For colorScale: minColor, midColor (optional), maxColor (hex strings) */
  colorScaleColors?: { min: string; mid?: string; max: string };
  /** For dataBar: fill color */
  dataBarColor?: string;
}
