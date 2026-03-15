import { WebviewState } from '../state/WebviewState';
import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { messageBridge } from '../state/MessageBridge';
import type { DataValidation } from '../../shared/types';
import { colToLetter, letterToCol } from '../../shared/types';
import { getElement } from './domUtils';

export class DvDialog {
  private dialog: HTMLElement;
  private dropdown: HTMLElement;
  private validations: DataValidation[] = [];

  constructor(
    private state: WebviewState,
    private renderer: CanvasRenderer,
  ) {
    this.dialog = getElement('dv-dialog');
    this.dropdown = getElement('dv-dropdown');
    this.setup();
    this.setupDropdownDismiss();
  }

  open(): void {
    this.dialog.classList.add('open');
    const sc = Math.min(this.state.selectedCol, this.state.selectionEndCol);
    const ec = Math.max(this.state.selectedCol, this.state.selectionEndCol);
    const sr = Math.min(this.state.selectedRow, this.state.selectionEndRow);
    const er = Math.max(this.state.selectedRow, this.state.selectionEndRow);
    getElement<HTMLInputElement>('dv-range').value =
      `${colToLetter(sc)}${sr + 1}:${colToLetter(ec)}${er + 1}`;
    this.renderList();
  }

  setValidations(validations: DataValidation[]): void {
    this.validations = validations;
    if (this.dialog.classList.contains('open')) this.renderList();
  }

  /** Show inline dropdown for list validation at a cell */
  showCellDropdown(col: number, row: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    const dv = this.validations.find(
      (v) =>
        v.type === 'list' &&
        v.range.sheet === sheet.name &&
        col >= v.range.startCol &&
        col <= v.range.endCol &&
        row >= v.range.startRow &&
        row <= v.range.endRow,
    );
    if (!dv) {
      this.dropdown.classList.remove('open');
      return;
    }

    const items = dv.listSource
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean);
    this.dropdown.innerHTML = '';
    for (const item of items) {
      const div = document.createElement('div');
      div.className = 'dv-option';
      div.textContent = item;
      div.addEventListener('mousedown', (e) => {
        e.preventDefault();
        messageBridge.postMessage({ type: 'editCell', sheet: sheet.name, col, row, value: item });
        this.dropdown.classList.remove('open');
      });
      this.dropdown.appendChild(div);
    }

    const { viewport, scrollManager } = this.renderer;
    const container = getElement('canvas-container');
    const containerRect = container.getBoundingClientRect();
    const cellX = 50 + viewport.colLeft(col) - scrollManager.scrollX;
    const cellY = 24 + viewport.rowTop(row) - scrollManager.scrollY;
    const cellH = viewport.rowHeight(row);
    this.dropdown.style.position = 'fixed';
    this.dropdown.style.left = `${containerRect.left + cellX}px`;
    this.dropdown.style.top = `${containerRect.top + cellY + cellH}px`;
    this.dropdown.style.minWidth = `${viewport.colWidth(col)}px`;
    this.dropdown.classList.add('open');
  }

  private renderList(): void {
    const list = getElement('dv-list');
    list.innerHTML = '';
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    const sheetValidations = this.validations.filter((v) => v.range.sheet === sheet.name);
    if (sheetValidations.length === 0) {
      list.innerHTML = '<div style="color:#888;padding:4px 0">No validations on this sheet</div>';
      return;
    }
    for (const dv of sheetValidations) {
      const r = dv.range;
      const rangeStr = `${colToLetter(r.startCol)}${r.startRow + 1}:${colToLetter(r.endCol)}${r.endRow + 1}`;
      const div = document.createElement('div');
      div.className = 'dv-item';
      div.innerHTML = `<span>${rangeStr} — ${dv.listSource}</span><span class="remove-dv" data-id="${dv.id}">\u2715</span>`;
      div.querySelector('.remove-dv')?.addEventListener('click', () => {
        const sh = this.state.activeSheet;
        if (!sh) return;
        messageBridge.postMessage({ type: 'removeDataValidation', sheet: sh.name, id: dv.id });
      });
      list.appendChild(div);
    }
  }

  private setup(): void {
    const dvTypeEl = getElement<HTMLSelectElement>('dv-type');
    const dvSourceLabel = getElement('dv-source-label');
    dvTypeEl.addEventListener('change', () => {
      dvSourceLabel.style.display = dvTypeEl.value === 'list' ? '' : 'none';
    });

    getElement('dv-apply').addEventListener('click', () => {
      const sheet = this.state.activeSheet;
      if (!sheet) return;
      const rangeStr = getElement<HTMLInputElement>('dv-range').value.trim();
      const type = dvTypeEl.value as 'list' | 'none';
      const source = getElement<HTMLInputElement>('dv-source').value.trim();

      const match = rangeStr.match(/^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/);
      if (!match) return;
      const startCol = letterToCol(match[1].toUpperCase());
      const startRow = parseInt(match[2], 10) - 1;
      const endCol = match[3] ? letterToCol(match[3].toUpperCase()) : startCol;
      const endRow = match[4] ? parseInt(match[4], 10) - 1 : startRow;

      const validation: DataValidation = {
        id: `dv_${Date.now()}`,
        range: { sheet: sheet.name, startCol, startRow, endCol, endRow },
        type,
        listSource: source,
      };
      messageBridge.postMessage({ type: 'setDataValidation', sheet: sheet.name, validation });
    });

    getElement('dv-close').addEventListener('click', () => {
      this.dialog.classList.remove('open');
    });
  }

  private setupDropdownDismiss(): void {
    document.addEventListener('click', (e) => {
      if (!this.dropdown.contains(e.target as Node)) {
        this.dropdown.classList.remove('open');
      }
    });
  }
}
