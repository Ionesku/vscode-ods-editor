/**
 * Shared types used by both the extension (src/) and the webview (webview/).
 * Webview code imports from here instead of directly from src/model/types.
 * Extension code can continue importing from src/model/types directly, or
 * from here — both resolve to the same declarations.
 */

export type {
  CellAddress,
  CellRange,
  CellValue,
  CellData,
  CellStyle,
  BorderStyle,
  SortSpec,
  FilterCriteria,
  DocumentMetadata,
  NamedRange,
  DataValidation,
} from '../src/model/types';

export {
  FormulaError,
  CellType,
  createEmptyCell,
  cellKey,
  parseCellKey,
  colToLetter,
  letterToCol,
} from '../src/model/types';
