import { WebviewState } from '../state/WebviewState';
import { messageBridge } from '../state/MessageBridge';
import { getElement } from './domUtils';

export class SheetTabs {
  private el: HTMLElement;
  private dragSrcIndex: number | null = null;

  constructor(private state: WebviewState) {
    this.el = getElement('sheet-tabs');
  }

  render(): void {
    this.el.innerHTML = '';

    this.state.sheets.forEach((sheet, index) => {
      const tab = document.createElement('div');
      tab.className = 'tab' + (index === this.state.activeSheetIndex ? ' active' : '');
      tab.textContent = sheet.name;
      tab.draggable = true;
      tab.dataset.index = String(index);

      tab.addEventListener('click', () => {
        messageBridge.postMessage({ type: 'switchSheet', index });
      });
      tab.addEventListener('dblclick', () => {
        const name = prompt('Rename sheet:', sheet.name);
        if (name) messageBridge.postMessage({ type: 'renameSheet', index, name });
      });
      tab.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        if (confirm(`Delete sheet "${sheet.name}"?`)) {
          messageBridge.postMessage({ type: 'removeSheet', index });
        }
      });

      tab.addEventListener('dragstart', (e) => {
        this.dragSrcIndex = index;
        tab.style.opacity = '0.5';
        if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
      });
      tab.addEventListener('dragend', () => {
        tab.style.opacity = '';
        this.el.querySelectorAll('.tab').forEach((t) => (t as HTMLElement).classList.remove('drag-over'));
      });
      tab.addEventListener('dragover', (e) => {
        e.preventDefault();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
        this.el.querySelectorAll('.tab').forEach((t) => (t as HTMLElement).classList.remove('drag-over'));
        tab.classList.add('drag-over');
      });
      tab.addEventListener('drop', (e) => {
        e.preventDefault();
        tab.classList.remove('drag-over');
        if (this.dragSrcIndex !== null && this.dragSrcIndex !== index) {
          messageBridge.postMessage({ type: 'reorderSheet', fromIndex: this.dragSrcIndex, toIndex: index });
          this.dragSrcIndex = null;
        }
      });

      this.el.appendChild(tab);
    });

    const addBtn = document.createElement('div');
    addBtn.className = 'add-sheet';
    addBtn.textContent = '+';
    addBtn.title = 'Add Sheet';
    addBtn.addEventListener('click', () => {
      messageBridge.postMessage({ type: 'addSheet', name: `Sheet${this.state.sheets.length + 1}` });
    });
    this.el.appendChild(addBtn);
  }
}
