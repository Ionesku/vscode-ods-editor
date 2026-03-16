import { WebviewState } from '../state/WebviewState';
import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { messageBridge } from '../state/MessageBridge';
import { getElement } from './domUtils';

export class FilterDropdown {
  private dropdown: HTMLElement;
  private valuesEl: HTMLElement;
  private selectAll: HTMLInputElement;

  private columnIndex = 0;
  private checkedValues = new Set<string>();

  constructor(
    private state: WebviewState,
    private renderer: CanvasRenderer,
  ) {
    this.dropdown = getElement('filter-dropdown');
    this.valuesEl = getElement('filter-values');
    this.selectAll = getElement<HTMLInputElement>('filter-select-all');
    this.setup();
  }

  show(column: number, values: string[]): void {
    this.columnIndex = column;
    const sheet = this.state.activeSheet;
    const existing = sheet?.activeFilters.get(column);

    this.valuesEl.innerHTML = '';
    this.checkedValues = new Set(existing ?? values);

    values.forEach((val) => {
      const label = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = this.checkedValues.has(val);
      cb.dataset.val = val;
      cb.addEventListener('change', () => {
        if (cb.checked) this.checkedValues.add(val);
        else this.checkedValues.delete(val);
        this.updateSelectAll(values.length);
      });
      label.appendChild(cb);
      label.appendChild(document.createTextNode(' ' + (val || '(empty)')));
      this.valuesEl.appendChild(label);
    });

    this.selectAll.checked = this.checkedValues.size === values.length;
    this.selectAll.indeterminate = this.checkedValues.size > 0 && this.checkedValues.size < values.length;
    this.selectAll.onchange = () => {
      const cbs = this.valuesEl.querySelectorAll('input[type=checkbox]') as NodeListOf<HTMLInputElement>;
      cbs.forEach((cb) => {
        cb.checked = this.selectAll.checked;
        const val = cb.dataset.val ?? '';
        if (this.selectAll.checked) this.checkedValues.add(val);
        else this.checkedValues.delete(val);
      });
    };

    const vp = this.renderer.viewport;
    const scrollX = this.renderer.scrollManager.scrollX;
    const colX = 50 + vp.colLeft(column) - scrollX;
    this.dropdown.style.left = colX + 'px';
    this.dropdown.style.top = '56px';
    this.dropdown.classList.add('open');
  }

  private setup(): void {
    getElement('filter-apply').addEventListener('click', () => {
      const sheet = this.state.activeSheet;
      if (!sheet) return;
      messageBridge.postMessage({
        type: 'setFilter',
        sheet: sheet.name,
        column: this.columnIndex,
        criteria: { column: this.columnIndex, values: Array.from(this.checkedValues) },
      });
      this.dropdown.classList.remove('open');
    });

    getElement('filter-clear').addEventListener('click', () => {
      const sheet = this.state.activeSheet;
      if (!sheet) return;
      messageBridge.postMessage({ type: 'setFilter', sheet: sheet.name, column: this.columnIndex, criteria: null });
      this.dropdown.classList.remove('open');
    });

    getElement('filter-close').addEventListener('click', () => {
      this.dropdown.classList.remove('open');
    });
  }

  private updateSelectAll(total: number): void {
    this.selectAll.checked = this.checkedValues.size === total;
    this.selectAll.indeterminate = this.checkedValues.size > 0 && this.checkedValues.size < total;
  }
}
