import { WebviewState } from '../state/WebviewState';
import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { ROW_HEADER_WIDTH, COL_HEADER_HEIGHT } from '../renderer/ViewportCalculator';

export class CellEditor {
  private textarea: HTMLTextAreaElement;
  private state: WebviewState;
  private renderer: CanvasRenderer;
  private onCommit: (value: string) => void;
  private onCancel: () => void;

  constructor(
    textarea: HTMLTextAreaElement,
    state: WebviewState,
    renderer: CanvasRenderer,
    onCommit: (value: string) => void,
    onCancel: () => void,
  ) {
    this.textarea = textarea;
    this.state = state;
    this.renderer = renderer;
    this.onCommit = onCommit;
    this.onCancel = onCancel;

    this.textarea.addEventListener('keydown', (e) => this.handleKeyDown(e));
    this.textarea.addEventListener('blur', () => {
      if (this.state.isEditing) {
        this.commit();
      }
    });
  }

  startEditing(initialValue?: string): void {
    const col = this.state.selectedCol;
    const row = this.state.selectedRow;
    const viewport = this.renderer.viewport;
    const scroll = this.renderer.scrollManager;

    const x = ROW_HEADER_WIDTH + viewport.colLeft(col) - scroll.scrollX;
    const y = COL_HEADER_HEIGHT + viewport.rowTop(row) - scroll.scrollY;
    const w = viewport.colWidth(col);
    const h = viewport.rowHeight(row);

    this.textarea.style.display = 'block';
    this.textarea.style.left = x + 'px';
    this.textarea.style.top = y + 'px';
    this.textarea.style.width = Math.max(w, 80) + 'px';
    this.textarea.style.height = h + 'px';

    if (initialValue !== undefined) {
      this.textarea.value = initialValue;
    } else {
      const cell = this.state.getCell(col, row);
      this.textarea.value = cell.formula ? '=' + cell.formula : String(cell.rawValue ?? '');
    }

    this.state.isEditing = true;
    this.state.editValue = this.textarea.value;
    this.textarea.focus();
    this.textarea.select();
  }

  stopEditing(): void {
    this.textarea.style.display = 'none';
    this.textarea.value = '';
    this.state.isEditing = false;
  }

  commit(): void {
    const value = this.textarea.value;
    this.stopEditing();
    this.onCommit(value);
  }

  cancel(): void {
    this.stopEditing();
    this.onCancel();
  }

  updateFormulaBar(formulaInput: HTMLInputElement): void {
    if (this.state.isEditing) {
      formulaInput.value = this.textarea.value;
    }
  }

  private handleKeyDown(e: KeyboardEvent): void {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      this.commit();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      this.cancel();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      this.commit();
    }
  }
}
