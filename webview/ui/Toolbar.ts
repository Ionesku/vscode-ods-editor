import { WebviewState } from '../state/WebviewState';
import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { messageBridge } from '../state/MessageBridge';
import { getElement } from './domUtils';
import type { CondFmtDialog } from './CondFmtDialog';
import type { NamedRangesDialog } from './NamedRangesDialog';
import type { DvDialog } from './DvDialog';
import type { FilterDropdown } from './FilterDropdown';

export class Toolbar {
  constructor(
    private state: WebviewState,
    private renderer: CanvasRenderer,
    private condFmtDialog: CondFmtDialog,
    private namedRangesDialog: NamedRangesDialog,
    private dvDialog: DvDialog,
    private filterDropdown: FilterDropdown,
  ) {
    this.setup();
  }

  private setup(): void {
    const btn = (id: string) => getElement(id);

    btn('btn-bold').addEventListener('click', () => this.applyStyle({ bold: true }));
    btn('btn-italic').addEventListener('click', () => this.applyStyle({ italic: true }));
    btn('btn-underline').addEventListener('click', () => this.applyStyle({ underline: true }));
    btn('btn-strikethrough').addEventListener('click', () => this.applyStyle({ strikethrough: true }));

    (btn('btn-text-color') as HTMLInputElement).addEventListener('input', (e) => {
      this.applyStyle({ textColor: (e.target as HTMLInputElement).value });
    });
    (btn('btn-bg-color') as HTMLInputElement).addEventListener('input', (e) => {
      this.applyStyle({ backgroundColor: (e.target as HTMLInputElement).value });
    });

    btn('btn-align-left').addEventListener('click', () => this.applyStyle({ horizontalAlign: 'left' }));
    btn('btn-align-center').addEventListener('click', () => this.applyStyle({ horizontalAlign: 'center' }));
    btn('btn-align-right').addEventListener('click', () => this.applyStyle({ horizontalAlign: 'right' }));

    this.setupDropdown('btn-borders', 'border-menu', (target) => {
      const type = target.dataset.border;
      if (type) this.applyBorder(type);
    });

    this.setupDropdown('btn-numfmt', 'numfmt-menu', (target) => {
      const fmt = target.dataset.fmt;
      if (!fmt) return;
      if (fmt === 'custom') {
        const customRow = getElement('custom-fmt-row');
        const customInput = getElement<HTMLInputElement>('custom-fmt-input');
        customRow.classList.toggle('visible');
        if (customRow.classList.contains('visible')) customInput.focus();
        return true; // keep menu open
      }
      this.applyNumberFormat(fmt);
    });

    getElement('custom-fmt-apply').addEventListener('click', () => {
      const input = getElement<HTMLInputElement>('custom-fmt-input');
      const fmt = input.value.trim();
      if (fmt) this.applyStyle({ numberFormat: fmt });
      getElement('numfmt-menu').classList.remove('open');
      getElement('custom-fmt-row').classList.remove('visible');
      input.value = '';
    });

    getElement<HTMLInputElement>('custom-fmt-input').addEventListener('input', (e) => {
      const fmt = (e.target as HTMLInputElement).value.trim();
      const preview = getElement('custom-fmt-preview');
      if (!fmt) { preview.textContent = ''; return; }
      const cell = this.state.getCell(this.state.selectedCol, this.state.selectedRow);
      const val = cell?.computedValue ?? cell?.rawValue;
      if (typeof val === 'number') {
        try {
          preview.textContent = this.state.formatForPreview(val, fmt);
        } catch {
          preview.textContent = '';
        }
      } else {
        preview.textContent = '';
      }
    });

    btn('btn-wrap').addEventListener('click', () => this.applyStyle({ wrapText: true }));
    btn('btn-freeze').addEventListener('click', () => this.toggleFreeze());
    btn('btn-sort-asc').addEventListener('click', () => this.sortSelection(true));
    btn('btn-sort-desc').addEventListener('click', () => this.sortSelection(false));
    btn('btn-filter').addEventListener('click', () => this.triggerFilter());
    btn('btn-cond-fmt').addEventListener('click', () => this.condFmtDialog.open());
    btn('btn-named-ranges').addEventListener('click', () => this.namedRangesDialog.open());
    btn('btn-data-validation').addEventListener('click', () => this.dvDialog.open());

    // Close dropdowns on outside click
    document.addEventListener('click', () => {
      document.querySelectorAll('.border-dropdown.open, .fmt-dropdown.open')
        .forEach((el) => el.classList.remove('open'));
    });
  }

  private setupDropdown(
    btnId: string,
    menuId: string,
    onSelect: (target: HTMLElement) => boolean | void,
  ): void {
    const btnEl = getElement(btnId);
    const menuEl = getElement(menuId);
    btnEl.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.border-dropdown.open, .fmt-dropdown.open').forEach((el) => {
        if (el !== menuEl) el.classList.remove('open');
      });
      menuEl.classList.toggle('open');
    });
    menuEl.addEventListener('click', (e) => {
      const target = (e.target as HTMLElement).closest('button') as HTMLElement;
      if (!target) return;
      e.stopPropagation();
      const keepOpen = onSelect(target);
      if (!keepOpen) menuEl.classList.remove('open');
    });
  }

  private applyStyle(style: Record<string, unknown>): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range) return;
    messageBridge.postMessage({ type: 'setStyle', sheet: sheet.name, range, style });
  }

  private applyBorder(type: string): void {
    const border = { color: '#000000', style: 'solid' as const, width: 'thin' as const };
    switch (type) {
      case 'all':
        this.applyStyle({ borderTop: border, borderBottom: border, borderLeft: border, borderRight: border });
        break;
      case 'outer':
        this.applyOuterBorders(border);
        break;
      case 'none':
        this.applyStyle({ borderTop: undefined, borderBottom: undefined, borderLeft: undefined, borderRight: undefined });
        break;
      case 'bottom': this.applyStyle({ borderBottom: border }); break;
      case 'top': this.applyStyle({ borderTop: border }); break;
      case 'left': this.applyStyle({ borderLeft: border }); break;
      case 'right': this.applyStyle({ borderRight: border }); break;
    }
  }

  private applyOuterBorders(border: { color: string; style: string; width: string }): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range) return;
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const style: Record<string, unknown> = {};
        if (r === range.startRow) style.borderTop = border;
        if (r === range.endRow) style.borderBottom = border;
        if (c === range.startCol) style.borderLeft = border;
        if (c === range.endCol) style.borderRight = border;
        if (Object.keys(style).length > 0) {
          messageBridge.postMessage({
            type: 'setStyle',
            sheet: sheet.name,
            range: { sheet: sheet.name, startCol: c, startRow: r, endCol: c, endRow: r },
            style,
          });
        }
      }
    }
  }

  private applyNumberFormat(fmt: string): void {
    const formats: Record<string, string> = {
      general: '',
      number: '#,##0.00',
      currency: '$#,##0.00',
      percent: '0.00%',
      date: 'yyyy-mm-dd',
      int: '#,##0',
      sci: '0.00E+00',
    };
    this.applyStyle({ numberFormat: formats[fmt] ?? '' });
  }

  private toggleFreeze(): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    if (sheet.frozenRows > 0 || sheet.frozenCols > 0) {
      messageBridge.postMessage({ type: 'freezePanes', sheet: sheet.name, frozenRows: 0, frozenCols: 0 });
    } else {
      messageBridge.postMessage({
        type: 'freezePanes',
        sheet: sheet.name,
        frozenRows: this.state.selectedRow,
        frozenCols: this.state.selectedCol,
      });
    }
  }

  private sortSelection(ascending: boolean): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range) return;
    messageBridge.postMessage({ type: 'sort', sheet: sheet.name, range, column: this.state.selectedCol, ascending });
  }

  private triggerFilter(): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    messageBridge.postMessage({ type: 'getColumnValues', sheet: sheet.name, column: this.state.selectedCol });
  }
}
