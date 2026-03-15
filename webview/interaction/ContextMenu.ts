import { WebviewState } from '../state/WebviewState';
import { messageBridge } from '../state/MessageBridge';
import { CanvasRenderer } from '../renderer/CanvasRenderer';

interface MenuItem {
  label: string;
  action: () => void;
  separator?: boolean;
}

export interface ClipboardActions {
  pasteValuesOnly(): void;
  pasteFormatsOnly(): void;
  hasClipboard(): boolean;
}

export class ContextMenu {
  private el: HTMLElement;
  private state: WebviewState;
  private renderer: CanvasRenderer;
  private clipboard: ClipboardActions | null = null;

  constructor(
    el: HTMLElement,
    state: WebviewState,
    renderer: CanvasRenderer,
    clipboard?: ClipboardActions,
  ) {
    this.el = el;
    this.state = state;
    this.renderer = renderer;
    this.clipboard = clipboard ?? null;

    // Close on click outside
    document.addEventListener('click', () => this.hide());
    document.addEventListener('contextmenu', (e) => {
      // Only show if right-click is on canvas
      const canvas = document.getElementById('spreadsheet-canvas');
      if (e.target === canvas) {
        e.preventDefault();
        this.show(e.clientX, e.clientY);
      }
    });
  }

  private show(x: number, y: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    const range = this.state.selectionRange;
    if (!range) return;

    const items: MenuItem[] = [
      { label: 'Insert Row Above', action: () => this.insertRows(range.startRow) },
      { label: 'Insert Row Below', action: () => this.insertRows(range.endRow + 1) },
      { label: 'Insert Column Left', action: () => this.insertCols(range.startCol) },
      { label: 'Insert Column Right', action: () => this.insertCols(range.endCol + 1) },
      { label: '', action: () => {}, separator: true },
      {
        label: 'Delete Rows',
        action: () => this.deleteRows(range.startRow, range.endRow - range.startRow + 1),
      },
      {
        label: 'Delete Columns',
        action: () => this.deleteCols(range.startCol, range.endCol - range.startCol + 1),
      },
      { label: '', action: () => {}, separator: true },
      { label: 'Sort Ascending', action: () => this.sort(true) },
      { label: 'Sort Descending', action: () => this.sort(false) },
      { label: '', action: () => {}, separator: true },
      { label: 'Merge Cells', action: () => this.merge() },
      { label: 'Unmerge Cells', action: () => this.unmerge() },
    ];

    if (this.clipboard?.hasClipboard()) {
      items.push({ label: '', action: () => {}, separator: true });
      items.push({
        label: 'Paste Values Only  (Ctrl+Shift+V)',
        action: () => this.clipboard!.pasteValuesOnly(),
      });
      items.push({ label: 'Paste Formats Only', action: () => this.clipboard!.pasteFormatsOnly() });
    }

    this.el.innerHTML = '';
    for (const item of items) {
      if (item.separator) {
        const sep = document.createElement('div');
        sep.className = 'menu-separator';
        this.el.appendChild(sep);
      } else {
        const div = document.createElement('div');
        div.className = 'menu-item';
        div.textContent = item.label;
        div.addEventListener('click', (e) => {
          e.stopPropagation();
          item.action();
          this.hide();
        });
        this.el.appendChild(div);
      }
    }

    this.el.style.left = x + 'px';
    this.el.style.top = y + 'px';
    this.el.style.display = 'block';
  }

  private hide(): void {
    this.el.style.display = 'none';
  }

  private insertRows(at: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({ type: 'insertRows', sheet: sheet.name, at, count: 1 });
  }

  private insertCols(at: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({ type: 'insertColumns', sheet: sheet.name, at, count: 1 });
  }

  private deleteRows(at: number, count: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({ type: 'deleteRows', sheet: sheet.name, at, count });
  }

  private deleteCols(at: number, count: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({ type: 'deleteColumns', sheet: sheet.name, at, count });
  }

  private sort(ascending: boolean): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range) return;
    messageBridge.postMessage({
      type: 'sort',
      sheet: sheet.name,
      range,
      column: this.state.selectedCol,
      ascending,
    });
  }

  private merge(): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range) return;
    messageBridge.postMessage({ type: 'mergeCells', sheet: sheet.name, range });
  }

  private unmerge(): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range) return;
    messageBridge.postMessage({ type: 'unmergeCells', sheet: sheet.name, range });
  }
}
