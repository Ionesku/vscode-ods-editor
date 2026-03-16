import { WebviewState } from '../state/WebviewState';
import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { SelectionManager } from '../interaction/SelectionManager';
import { messageBridge } from '../state/MessageBridge';
import { getElement } from './domUtils';

export class FindBar {
  private bar: HTMLElement;
  private findInput: HTMLInputElement;
  private replaceInput: HTMLInputElement;
  private countEl: HTMLElement;
  private canvas: HTMLCanvasElement;

  private matches: Array<{ col: number; row: number }> = [];
  private matchIndex = -1;

  constructor(
    private state: WebviewState,
    private renderer: CanvasRenderer,
    private selection: SelectionManager,
  ) {
    this.bar = getElement('find-bar');
    this.findInput = getElement<HTMLInputElement>('find-input');
    this.replaceInput = getElement<HTMLInputElement>('replace-input');
    this.countEl = getElement('find-count');
    this.canvas = getElement<HTMLCanvasElement>('spreadsheet-canvas');
    this.setup();
  }

  open(): void {
    this.bar.classList.add('open');
    this.findInput.focus();
    this.findInput.select();
  }

  close(): void {
    this.bar.classList.remove('open');
    this.matches = [];
    this.matchIndex = -1;
    this.countEl.textContent = '';
    this.canvas.focus();
    this.renderer.markDirty();
  }

  isOpen(): boolean {
    return this.bar.classList.contains('open');
  }

  private setup(): void {
    this.findInput.addEventListener('input', () => this.performFind());
    this.findInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        e.shiftKey ? this.findPrev() : this.findNext();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      }
    });
    this.replaceInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.replaceOne();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.close();
      }
    });
    getElement('find-next').addEventListener('click', () => this.findNext());
    getElement('find-prev').addEventListener('click', () => this.findPrev());
    getElement('replace-one').addEventListener('click', () => this.replaceOne());
    getElement('replace-all').addEventListener('click', () => this.replaceAll());
    getElement('find-close').addEventListener('click', () => this.close());
  }

  private performFind(): void {
    const query = this.findInput.value.toLowerCase();
    this.matches = [];
    this.matchIndex = -1;

    if (!query || !this.state.activeSheet) {
      this.countEl.textContent = '';
      this.renderer.markDirty();
      return;
    }

    const sheet = this.state.activeSheet;
    for (const [key, cell] of sheet.cells) {
      const val = cell.computedValue ?? cell.rawValue;
      if (val !== null && String(val).toLowerCase().includes(query)) {
        const idx = key.indexOf(',');
        this.matches.push({
          col: parseInt(key.substring(0, idx), 10),
          row: parseInt(key.substring(idx + 1), 10),
        });
      }
    }

    this.matches.sort((a, b) => a.row - b.row || a.col - b.col);

    if (this.matches.length > 0) {
      this.matchIndex = 0;
      this.navigateToMatch();
    }
    this.updateCount();
    this.renderer.markDirty();
  }

  private navigateToMatch(): void {
    if (this.matchIndex < 0 || this.matchIndex >= this.matches.length) return;
    const m = this.matches[this.matchIndex];
    this.selection.selectCell(m.col, m.row);
    this.renderer.scrollManager.scrollToCell(
      m.col, m.row,
      this.state.activeSheet?.frozenCols ?? 0,
      this.state.activeSheet?.frozenRows ?? 0,
    );
    this.renderer.markDirty();
    this.updateCount();
  }

  private findNext(): void {
    if (this.matches.length === 0) return;
    this.matchIndex = (this.matchIndex + 1) % this.matches.length;
    this.navigateToMatch();
  }

  private findPrev(): void {
    if (this.matches.length === 0) return;
    this.matchIndex = (this.matchIndex - 1 + this.matches.length) % this.matches.length;
    this.navigateToMatch();
  }

  private replaceOne(): void {
    if (this.matchIndex < 0 || this.matchIndex >= this.matches.length) return;
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    const m = this.matches[this.matchIndex];
    messageBridge.postMessage({ type: 'editCell', sheet: sheet.name, col: m.col, row: m.row, value: this.replaceInput.value });
    this.matches.splice(this.matchIndex, 1);
    if (this.matchIndex >= this.matches.length) this.matchIndex = 0;
    if (this.matches.length > 0) this.navigateToMatch();
    this.updateCount();
  }

  private replaceAll(): void {
    const sheet = this.state.activeSheet;
    if (!sheet || this.matches.length === 0) return;
    const replacement = this.replaceInput.value;
    const query = this.findInput.value.toLowerCase();

    for (const m of this.matches) {
      const cell = this.state.getCell(m.col, m.row);
      const val = cell.computedValue ?? cell.rawValue;
      if (val === null) continue;
      const newVal = String(val).replace(new RegExp(escapeRegex(query), 'gi'), replacement);
      messageBridge.postMessage({ type: 'editCell', sheet: sheet.name, col: m.col, row: m.row, value: newVal });
    }
    this.matches = [];
    this.matchIndex = -1;
    this.updateCount();
  }

  private updateCount(): void {
    if (this.matches.length === 0) {
      this.countEl.textContent = this.findInput.value ? '0 results' : '';
    } else {
      this.countEl.textContent = `${this.matchIndex + 1} of ${this.matches.length}`;
    }
  }
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
