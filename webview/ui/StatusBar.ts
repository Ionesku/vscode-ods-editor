import { WebviewState } from '../state/WebviewState';
import { getElement } from './domUtils';

export class StatusBar {
  private el: HTMLElement;

  constructor(private state: WebviewState) {
    this.el = getElement('status-bar');
  }

  update(): void {
    const range = this.state.selectionRange;
    if (!range) {
      this.el.textContent = '';
      return;
    }

    const isSingle = range.startCol === range.endCol && range.startRow === range.endRow;
    if (isSingle) {
      this.el.textContent = '';
      return;
    }

    let sum = 0;
    let count = 0;
    let numCount = 0;

    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const cell = this.state.getCell(c, r);
        const val = cell.computedValue ?? cell.rawValue;
        if (val !== null && val !== '') {
          count++;
          const num = typeof val === 'number' ? val : Number(val);
          if (!isNaN(num) && typeof val !== 'boolean') {
            sum += num;
            numCount++;
          }
        }
      }
    }

    const parts: string[] = [];
    if (numCount > 0) {
      parts.push(`Sum: ${this.formatNum(sum)}`);
      parts.push(`Avg: ${this.formatNum(sum / numCount)}`);
    }
    parts.push(`Count: ${count}`);
    this.el.textContent = parts.join('   ');
  }

  private formatNum(n: number): string {
    if (Number.isInteger(n)) return String(n);
    return n.toFixed(2);
  }
}
