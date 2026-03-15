import { WebviewState } from '../state/WebviewState';
import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { SelectionManager } from '../interaction/SelectionManager';
import { messageBridge } from '../state/MessageBridge';
import { colToLetter, letterToCol } from '../../shared/types';
import { getElement } from './domUtils';
import { FormulaAutocomplete } from './FormulaAutocomplete';

export class FormulaBar {
  private cellAddressInput: HTMLInputElement;
  private formulaInput: HTMLInputElement;
  private canvas: HTMLCanvasElement;
  private autocomplete: FormulaAutocomplete;

  constructor(
    private state: WebviewState,
    private renderer: CanvasRenderer,
    private selection: SelectionManager,
  ) {
    this.cellAddressInput = getElement<HTMLInputElement>('cell-address');
    this.formulaInput = getElement<HTMLInputElement>('formula-input');
    this.canvas = getElement<HTMLCanvasElement>('spreadsheet-canvas');
    this.autocomplete = new FormulaAutocomplete(this.formulaInput, state);
    this.setup();
  }

  updateAddress(): void {
    this.cellAddressInput.value =
      colToLetter(this.state.selectedCol) + (this.state.selectedRow + 1);
  }

  updateFormula(): void {
    const cell = this.state.getCell(this.state.selectedCol, this.state.selectedRow);
    if (cell.formula) {
      this.formulaInput.value = '=' + cell.formula;
    } else {
      this.formulaInput.value = cell.rawValue !== null ? String(cell.rawValue) : '';
    }
  }

  private setup(): void {
    this.formulaInput.addEventListener('keydown', (e) => {
      // Let autocomplete handle arrow/tab/enter/escape when it's visible
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') return;
      if (e.key === 'Enter') {
        e.preventDefault();
        this.autocomplete.hide();
        const sheet = this.state.activeSheet;
        if (!sheet) return;
        messageBridge.postMessage({
          type: 'editCell',
          sheet: sheet.name,
          col: this.state.selectedCol,
          row: this.state.selectedRow,
          value: this.formulaInput.value,
        });
        this.canvas.focus();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.autocomplete.hide();
        this.updateFormula();
        this.canvas.focus();
      }
    });

    this.cellAddressInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.navigateToAddress(this.cellAddressInput.value.trim().toUpperCase());
        this.canvas.focus();
      } else if (e.key === 'Escape') {
        this.updateAddress();
        this.canvas.focus();
      }
    });
  }

  private navigateToAddress(raw: string): void {
    const rangeMatch = raw.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    const cellMatch = raw.match(/^([A-Z]+)(\d+)$/);

    if (rangeMatch) {
      const c1 = this.parseLetter(rangeMatch[1]);
      const r1 = parseInt(rangeMatch[2], 10) - 1;
      const c2 = this.parseLetter(rangeMatch[3]);
      const r2 = parseInt(rangeMatch[4], 10) - 1;
      if (c1 >= 0 && r1 >= 0 && c2 >= 0 && r2 >= 0) {
        this.selection.selectCell(c1, r1);
        this.selection.selectCell(c2, r2, true);
        this.renderer.scrollManager.scrollToCell(
          c1, r1,
          this.state.activeSheet?.frozenCols ?? 0,
          this.state.activeSheet?.frozenRows ?? 0,
        );
        this.renderer.markDirty();
      }
    } else if (cellMatch) {
      const c = this.parseLetter(cellMatch[1]);
      const r = parseInt(cellMatch[2], 10) - 1;
      if (c >= 0 && r >= 0) {
        this.selection.selectCell(c, r);
        this.renderer.scrollManager.scrollToCell(
          c, r,
          this.state.activeSheet?.frozenCols ?? 0,
          this.state.activeSheet?.frozenRows ?? 0,
        );
        this.renderer.markDirty();
      }
    }
  }

  private parseLetter(letters: string): number {
    return letterToCol(letters);
  }
}
