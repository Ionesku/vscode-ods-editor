import { WebviewState } from '../state/WebviewState';
import { colToLetter } from '../../shared/types';

export class SelectionManager {
  private state: WebviewState;
  private onSelectionChange: () => void;

  constructor(state: WebviewState, onSelectionChange: () => void) {
    this.state = state;
    this.onSelectionChange = onSelectionChange;
  }

  selectCell(col: number, row: number, extend = false): void {
    if (extend) {
      this.state.selectionEndCol = col;
      this.state.selectionEndRow = row;
    } else {
      // Check if clicking on a merged cell — expand selection to full merge area
      const cell = this.state.getCell(col, row);
      if (cell.mergedInto) {
        // Clicked on a cell merged into another — select the anchor
        const anchor = cell.mergedInto;
        const anchorCell = this.state.getCell(anchor.col, anchor.row);
        this.state.selectedCol = anchor.col;
        this.state.selectedRow = anchor.row;
        this.state.selectionEndCol = anchor.col + anchorCell.mergeColSpan - 1;
        this.state.selectionEndRow = anchor.row + anchorCell.mergeRowSpan - 1;
      } else if (cell.mergeColSpan > 1 || cell.mergeRowSpan > 1) {
        // Clicked on the anchor of a merged cell
        this.state.selectedCol = col;
        this.state.selectedRow = row;
        this.state.selectionEndCol = col + cell.mergeColSpan - 1;
        this.state.selectionEndRow = row + cell.mergeRowSpan - 1;
      } else {
        this.state.selectedCol = col;
        this.state.selectedRow = row;
        this.state.selectionEndCol = col;
        this.state.selectionEndRow = row;
      }
    }
    this.onSelectionChange();
  }

  selectColumn(col: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    this.state.selectedCol = col;
    this.state.selectedRow = 0;
    this.state.selectionEndCol = col;
    this.state.selectionEndRow = sheet.rowCount - 1;
    this.onSelectionChange();
  }

  selectRow(row: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    this.state.selectedCol = 0;
    this.state.selectedRow = row;
    this.state.selectionEndCol = sheet.columnCount - 1;
    this.state.selectionEndRow = row;
    this.onSelectionChange();
  }

  moveSelection(dCol: number, dRow: number, extend = false): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;

    if (extend) {
      this.state.selectionEndCol = Math.max(
        0,
        Math.min(sheet.columnCount - 1, this.state.selectionEndCol + dCol),
      );
      this.state.selectionEndRow = Math.max(
        0,
        Math.min(sheet.rowCount - 1, this.state.selectionEndRow + dRow),
      );
    } else {
      this.state.selectedCol = Math.max(
        0,
        Math.min(sheet.columnCount - 1, this.state.selectedCol + dCol),
      );
      this.state.selectedRow = Math.max(
        0,
        Math.min(sheet.rowCount - 1, this.state.selectedRow + dRow),
      );
      this.state.selectionEndCol = this.state.selectedCol;
      this.state.selectionEndRow = this.state.selectedRow;
    }
    this.onSelectionChange();
  }

  selectAll(): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    this.state.selectedCol = 0;
    this.state.selectedRow = 0;
    this.state.selectionEndCol = sheet.columnCount - 1;
    this.state.selectionEndRow = sheet.rowCount - 1;
    this.onSelectionChange();
  }

  getCellAddress(): string {
    return colToLetter(this.state.selectedCol) + (this.state.selectedRow + 1);
  }
}
