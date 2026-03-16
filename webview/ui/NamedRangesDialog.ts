import { WebviewState } from '../state/WebviewState';
import { messageBridge } from '../state/MessageBridge';
import { colToLetter, letterToCol } from '../../shared/types';
import { getElement } from './domUtils';

type NamedRangeEntry = {
  name: string;
  sheet: string;
  range: { startCol: number; startRow: number; endCol: number; endRow: number };
};

export class NamedRangesDialog {
  private dialog: HTMLElement;
  private namedRanges: NamedRangeEntry[] = [];

  constructor(private state: WebviewState) {
    this.dialog = getElement('named-ranges-dialog');
    this.setup();
  }

  open(): void {
    this.dialog.classList.add('open');
    this.renderList();
    const range = this.state.selectionRange;
    if (range) {
      const addr = `${colToLetter(range.startCol)}${range.startRow + 1}:${colToLetter(range.endCol)}${range.endRow + 1}`;
      getElement<HTMLInputElement>('nr-range').value = addr;
    }
  }

  setNamedRanges(ranges: NamedRangeEntry[]): void {
    this.namedRanges = ranges;
    if (this.dialog.classList.contains('open')) {
      this.renderList();
    }
  }

  private renderList(): void {
    const list = getElement('nr-list');
    list.innerHTML = '';
    for (const nr of this.namedRanges) {
      const div = document.createElement('div');
      div.className = 'nr-item';

      const span = document.createElement('span');
      span.textContent = `${nr.name}  →  ${nr.sheet}!${colToLetter(nr.range.startCol)}${nr.range.startRow + 1}:${colToLetter(nr.range.endCol)}${nr.range.endRow + 1}`;
      div.appendChild(span);

      const rm = document.createElement('span');
      rm.className = 'remove-nr';
      rm.textContent = '\u00d7';
      rm.addEventListener('click', () => {
        messageBridge.postMessage({ type: 'deleteNamedRange', name: nr.name });
      });
      div.appendChild(rm);
      list.appendChild(div);
    }
  }

  private setup(): void {
    getElement('nr-define').addEventListener('click', () => {
      const sheet = this.state.activeSheet;
      if (!sheet) return;
      const name = getElement<HTMLInputElement>('nr-name').value.trim();
      const rangeStr = getElement<HTMLInputElement>('nr-range').value.trim().toUpperCase();
      if (!name) { alert('Enter a name'); return; }

      const match = rangeStr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (!match) { alert('Invalid range format. Use A1:B10'); return; }

      messageBridge.postMessage({
        type: 'defineNamedRange',
        name,
        sheet: sheet.name,
        range: {
          sheet: sheet.name,
          startCol: letterToCol(match[1]),
          startRow: parseInt(match[2]) - 1,
          endCol: letterToCol(match[3]),
          endRow: parseInt(match[4]) - 1,
        },
      });
      getElement<HTMLInputElement>('nr-name').value = '';
    });

    getElement('nr-close').addEventListener('click', () => {
      this.dialog.classList.remove('open');
    });
  }
}
