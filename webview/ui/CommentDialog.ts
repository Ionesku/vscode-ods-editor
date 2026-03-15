import { WebviewState } from '../state/WebviewState';
import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { messageBridge } from '../state/MessageBridge';
import { getElement } from './domUtils';

/**
 * Dialog for editing cell comments.
 * Also shows a hover tooltip when the mouse is over a cell with a comment.
 */
export class CommentDialog {
  private dialog: HTMLElement;
  private textarea: HTMLTextAreaElement;
  private tooltip: HTMLElement;
  private currentCol = 0;
  private currentRow = 0;

  constructor(
    private state: WebviewState,
    private renderer: CanvasRenderer,
    canvas: HTMLCanvasElement,
  ) {
    this.dialog = getElement('comment-dialog');
    this.textarea = getElement<HTMLTextAreaElement>('comment-text');
    this.tooltip = getElement('comment-tooltip');
    this.setup();
    this.setupTooltip(canvas);
  }

  open(col: number, row: number): void {
    this.currentCol = col;
    this.currentRow = row;
    const cell = this.state.getCell(col, row);
    this.textarea.value = cell.comment ?? '';
    this.dialog.classList.add('open');
    this.textarea.focus();
  }

  private setup(): void {
    getElement('comment-save').addEventListener('click', () => {
      const sheet = this.state.activeSheet;
      if (!sheet) return;
      messageBridge.postMessage({
        type: 'setComment',
        sheet: sheet.name,
        col: this.currentCol,
        row: this.currentRow,
        comment: this.textarea.value.trim(),
      });
      this.dialog.classList.remove('open');
    });

    getElement('comment-delete').addEventListener('click', () => {
      const sheet = this.state.activeSheet;
      if (!sheet) return;
      messageBridge.postMessage({
        type: 'setComment',
        sheet: sheet.name,
        col: this.currentCol,
        row: this.currentRow,
        comment: '',
      });
      this.dialog.classList.remove('open');
    });

    getElement('comment-cancel').addEventListener('click', () => {
      this.dialog.classList.remove('open');
    });

    this.textarea.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') this.dialog.classList.remove('open');
    });
  }

  private setupTooltip(canvas: HTMLCanvasElement): void {
    canvas.addEventListener('mousemove', (e) => {
      const rect = canvas.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      const hit = this.renderer.hitTest(x, y);
      if (!hit) { this.tooltip.style.display = 'none'; return; }

      const cell = this.state.getCell(hit.col, hit.row);
      if (!cell.comment) { this.tooltip.style.display = 'none'; return; }

      this.tooltip.textContent = cell.comment;
      this.tooltip.style.display = 'block';
      this.tooltip.style.left = (e.clientX + 12) + 'px';
      this.tooltip.style.top = (e.clientY + 12) + 'px';
    });

    canvas.addEventListener('mouseleave', () => {
      this.tooltip.style.display = 'none';
    });
  }
}
