import { WebviewState } from '../state/WebviewState';

/**
 * Shows a function-name autocomplete dropdown below the formula input.
 * Activates when the formula starts with "=" and the cursor is on a function name token.
 */
export class FormulaAutocomplete {
  private dropdown: HTMLElement;
  private selectedIndex = -1;
  private suggestions: string[] = [];
  private onPick: (name: string) => void = () => {};

  constructor(
    private input: HTMLInputElement,
    private state: WebviewState,
  ) {
    this.dropdown = this.createDropdown();
    document.body.appendChild(this.dropdown);
    this.bindEvents();
  }

  private createDropdown(): HTMLElement {
    const el = document.createElement('div');
    el.id = 'formula-autocomplete';
    el.style.cssText =
      'display:none;position:fixed;z-index:500;background:var(--vscode-editorSuggestWidget-background,#252526);' +
      'border:1px solid var(--vscode-editorSuggestWidget-border,#454545);' +
      'box-shadow:0 2px 8px rgba(0,0,0,0.4);max-height:180px;overflow-y:auto;min-width:160px;font-size:12px;';
    return el;
  }

  private bindEvents(): void {
    this.input.addEventListener('input', () => this.update());
    this.input.addEventListener('keydown', (e) => this.handleKey(e));
    this.input.addEventListener('blur', () => {
      // Slight delay so click on dropdown item registers first
      setTimeout(() => this.hide(), 150);
    });
  }

  private update(): void {
    const val = this.input.value;
    if (!val.startsWith('=')) { this.hide(); return; }

    // Find the function name token at the cursor position
    const cursor = this.input.selectionStart ?? val.length;
    const prefix = val.slice(1, cursor); // strip leading '='
    // Extract the last identifier (letters only, before any operator/paren)
    const match = prefix.match(/([A-Za-z][A-Za-z0-9]*)$/);
    if (!match) { this.hide(); return; }
    const token = match[1].toUpperCase();
    if (token.length < 1) { this.hide(); return; }

    this.suggestions = this.state.functionNames.filter((fn) => fn.startsWith(token));
    if (this.suggestions.length === 0 || (this.suggestions.length === 1 && this.suggestions[0] === token)) {
      this.hide(); return;
    }

    this.selectedIndex = 0;
    this.render();
    this.positionBelow();
    this.dropdown.style.display = 'block';
  }

  private render(): void {
    this.dropdown.innerHTML = '';
    this.suggestions.forEach((name, i) => {
      const item = document.createElement('div');
      item.textContent = name;
      item.style.cssText =
        'padding:3px 10px;cursor:pointer;white-space:nowrap;' +
        (i === this.selectedIndex
          ? 'background:var(--vscode-editorSuggestWidget-selectedBackground,#094771);color:var(--vscode-editorSuggestWidget-selectedForeground,#fff);'
          : '');
      item.addEventListener('mousedown', (e) => {
        e.preventDefault();
        this.accept(name);
      });
      this.dropdown.appendChild(item);
    });
  }

  private handleKey(e: KeyboardEvent): void {
    if (!this.isVisible()) return;
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      this.selectedIndex = Math.min(this.selectedIndex + 1, this.suggestions.length - 1);
      this.render();
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      this.selectedIndex = Math.max(this.selectedIndex - 1, 0);
      this.render();
    } else if (e.key === 'Tab' || e.key === 'Enter') {
      if (this.suggestions.length > 0 && this.selectedIndex >= 0) {
        e.preventDefault();
        this.accept(this.suggestions[this.selectedIndex]);
      }
    } else if (e.key === 'Escape') {
      this.hide();
    }
  }

  private accept(name: string): void {
    const val = this.input.value;
    const cursor = this.input.selectionStart ?? val.length;
    const before = val.slice(0, cursor);
    const after = val.slice(cursor);
    // Replace the last word (function token) with the accepted name
    const replaced = before.replace(/([A-Za-z][A-Za-z0-9]*)$/, name);
    this.input.value = replaced + '(' + after;
    // Position cursor inside the parentheses
    const newCursor = replaced.length + 1;
    this.input.setSelectionRange(newCursor, newCursor);
    this.hide();
    this.onPick(name);
  }

  private positionBelow(): void {
    const rect = this.input.getBoundingClientRect();
    this.dropdown.style.left = rect.left + 'px';
    this.dropdown.style.top = rect.bottom + 2 + 'px';
    this.dropdown.style.width = Math.max(rect.width, 160) + 'px';
  }

  private isVisible(): boolean {
    return this.dropdown.style.display !== 'none';
  }

  hide(): void {
    this.dropdown.style.display = 'none';
    this.suggestions = [];
    this.selectedIndex = -1;
  }

  setOnPick(cb: (name: string) => void): void {
    this.onPick = cb;
  }
}
