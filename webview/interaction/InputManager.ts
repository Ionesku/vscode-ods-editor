import { CanvasRenderer } from '../renderer/CanvasRenderer';
import { WebviewState } from '../state/WebviewState';
import { SelectionManager } from './SelectionManager';
import { CellEditor } from './CellEditor';
import { messageBridge } from '../state/MessageBridge';

export class InputManager {
  private canvas: HTMLCanvasElement;
  private state: WebviewState;
  private renderer: CanvasRenderer;
  private selection: SelectionManager;
  private editor: CellEditor;
  private isMouseDown = false;

  // Resize state
  private resizingCol: number | null = null;
  private resizingRow: number | null = null;
  private resizeStartX = 0;
  private resizeStartY = 0;
  private resizeOrigSize = 0;

  // Auto-fill state
  private isFillDragging = false;
  private fillDragEndCol = 0;
  private fillDragEndRow = 0;

  // Cell drag-to-move state
  private isMoveDragging = false;
  private moveDragTargetCol = 0;
  private moveDragTargetRow = 0;

  constructor(
    canvas: HTMLCanvasElement,
    state: WebviewState,
    renderer: CanvasRenderer,
    selection: SelectionManager,
    editor: CellEditor,
  ) {
    this.canvas = canvas;
    this.state = state;
    this.renderer = renderer;
    this.selection = selection;
    this.editor = editor;

    canvas.addEventListener('mousedown', (e) => this.handleMouseDown(e));
    canvas.addEventListener('mousemove', (e) => this.handleMouseMove(e));
    canvas.addEventListener('mouseup', () => this.handleMouseUp());
    canvas.addEventListener('dblclick', (e) => this.handleDblClick(e));
    document.addEventListener('keydown', (e) => this.handleKeyDown(e));

    // DOM clipboard events — работают в VSCode webview без разрешений
    document.addEventListener('copy', (e) => this.handleCopy(e));
    document.addEventListener('cut', (e) => this.handleCut(e));
    document.addEventListener('paste', (e) => this.handlePaste(e));
  }

  /** scrollToCell that automatically passes the current sheet's frozen pane info */
  private scrollToCell(col: number, row: number): void {
    const sheet = this.state.activeSheet;
    this.renderer.scrollManager.scrollToCell(
      col,
      row,
      sheet?.frozenCols ?? 0,
      sheet?.frozenRows ?? 0,
    );
  }

  private getCanvasCoords(e: MouseEvent): { x: number; y: number } {
    const rect = this.canvas.getBoundingClientRect();
    return { x: e.clientX - rect.left, y: e.clientY - rect.top };
  }

  private handleMouseDown(e: MouseEvent): void {
    // Ignore right-click (preserve selection for context menu)
    if (e.button === 2) return;

    const { x, y } = this.getCanvasCoords(e);

    // Check for fill handle drag
    if (this.renderer.hitTestFillHandle(x, y)) {
      const range = this.state.selectionRange;
      if (range) {
        this.isFillDragging = true;
        this.fillDragEndCol = range.endCol;
        this.fillDragEndRow = range.endRow;
        e.preventDefault();
        return;
      }
    }

    // Check for move drag: Alt + mousedown inside current selection body
    if (e.altKey) {
      const hit = this.renderer.hitTest(x, y);
      const range = this.state.selectionRange;
      if (hit && range &&
        hit.col >= range.startCol && hit.col <= range.endCol &&
        hit.row >= range.startRow && hit.row <= range.endRow) {
        this.isMoveDragging = true;
        this.moveDragTargetCol = range.startCol;
        this.moveDragTargetRow = range.startRow;
        this.canvas.style.cursor = 'move';
        e.preventDefault();
        return;
      }
    }

    // Check for column resize (must check before header click)
    const resizeCol = this.renderer.hitTestColResize(x, y);
    if (resizeCol !== null) {
      this.resizingCol = resizeCol;
      this.resizeStartX = e.clientX;
      this.resizeOrigSize = this.state.activeSheet?.columnWidths[resizeCol] ?? 80;
      e.preventDefault();
      return;
    }

    // Check for row resize
    const resizeRow = this.renderer.hitTestRowResize(x, y);
    if (resizeRow !== null) {
      this.resizingRow = resizeRow;
      this.resizeStartY = e.clientY;
      this.resizeOrigSize = this.state.activeSheet?.rowHeights[resizeRow] ?? 24;
      e.preventDefault();
      return;
    }

    // Check for column header click → select entire column
    const colHeader = this.renderer.hitTestColHeader(x, y);
    if (colHeader !== null) {
      if (this.state.isEditing) this.editor.commit();
      this.selection.selectColumn(colHeader);
      this.renderer.markDirty();
      return;
    }

    // Check for row header click → select entire row
    const rowHeader = this.renderer.hitTestRowHeader(x, y);
    if (rowHeader !== null) {
      if (this.state.isEditing) this.editor.commit();
      this.selection.selectRow(rowHeader);
      this.renderer.markDirty();
      return;
    }

    // Cell selection
    const hit = this.renderer.hitTest(x, y);
    if (hit) {
      if (this.state.isEditing) {
        this.editor.commit();
      }
      this.selection.selectCell(hit.col, hit.row, e.shiftKey);
      this.scrollToCell(this.state.selectedCol, this.state.selectedRow);
      this.renderer.markDirty();
      this.isMouseDown = true;
    }
  }

  private handleMouseMove(e: MouseEvent): void {
    const { x, y } = this.getCanvasCoords(e);

    // Handle move drag
    if (this.isMoveDragging) {
      const hit = this.renderer.hitTest(x, y);
      if (hit) {
        this.moveDragTargetCol = hit.col;
        this.moveDragTargetRow = hit.row;
        this.renderer.markSelectionDirty();
      }
      return;
    }

    // Handle fill drag
    if (this.isFillDragging) {
      const hit = this.renderer.hitTest(x, y);
      if (hit) {
        this.fillDragEndCol = hit.col;
        this.fillDragEndRow = hit.row;
        this.renderer.markDirty();
      }
      return;
    }

    // Handle column resize drag
    if (this.resizingCol !== null) {
      const delta = e.clientX - this.resizeStartX;
      const newWidth = Math.max(20, this.resizeOrigSize + delta);
      const sheet = this.state.activeSheet;
      if (sheet) {
        sheet.columnWidths[this.resizingCol] = newWidth;
        this.renderer.updateViewport();
        this.renderer.markDirty();
      }
      return;
    }

    // Handle row resize drag
    if (this.resizingRow !== null) {
      const delta = e.clientY - this.resizeStartY;
      const newHeight = Math.max(10, this.resizeOrigSize + delta);
      const sheet = this.state.activeSheet;
      if (sheet) {
        sheet.rowHeights[this.resizingRow] = newHeight;
        this.renderer.updateViewport();
        this.renderer.markDirty();
      }
      return;
    }

    // Update cursor for resize/fill handles
    const colResize = this.renderer.hitTestColResize(x, y);
    const rowResize = this.renderer.hitTestRowResize(x, y);
    const fillHandle = this.renderer.hitTestFillHandle(x, y);
    if (fillHandle) this.canvas.style.cursor = 'crosshair';
    else if (colResize !== null) this.canvas.style.cursor = 'col-resize';
    else if (rowResize !== null) this.canvas.style.cursor = 'row-resize';
    else this.canvas.style.cursor = 'cell';

    // Drag selection
    if (this.isMouseDown) {
      const hit = this.renderer.hitTest(x, y);
      if (hit) {
        this.selection.selectCell(hit.col, hit.row, true);
        this.renderer.markDirty();
      }
    }
  }

  private handleMouseUp(): void {
    // Handle move drag completion
    if (this.isMoveDragging) {
      this.isMoveDragging = false;
      this.canvas.style.cursor = 'cell';
      const sheet = this.state.activeSheet;
      const range = this.state.selectionRange;
      if (sheet && range) {
        const toCol = this.moveDragTargetCol;
        const toRow = this.moveDragTargetRow;
        if (toCol !== range.startCol || toRow !== range.startRow) {
          messageBridge.postMessage({
            type: 'moveRange',
            sheet: sheet.name,
            fromRange: range,
            toCol,
            toRow,
          });
        }
      }
      return;
    }

    // Handle fill drag completion
    if (this.isFillDragging) {
      this.completeFillDrag();
      this.isFillDragging = false;
      return;
    }

    // Commit resize
    if (this.resizingCol !== null) {
      const sheet = this.state.activeSheet;
      if (sheet) {
        messageBridge.postMessage({
          type: 'resizeColumn',
          sheet: sheet.name,
          col: this.resizingCol,
          width: sheet.columnWidths[this.resizingCol],
        });
      }
      this.resizingCol = null;
    }
    if (this.resizingRow !== null) {
      const sheet = this.state.activeSheet;
      if (sheet) {
        messageBridge.postMessage({
          type: 'resizeRow',
          sheet: sheet.name,
          row: this.resizingRow,
          height: sheet.rowHeights[this.resizingRow],
        });
      }
      this.resizingRow = null;
    }
    this.isMouseDown = false;
  }

  // ── Auto-fill logic ──────────────────────────────────────────────────────

  private completeFillDrag(): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range) return;

    const dragEndRow = this.fillDragEndRow;
    const dragEndCol = this.fillDragEndCol;

    // Determine fill direction: only fill down or right from selection
    if (dragEndRow > range.endRow) {
      // Fill down
      for (let c = range.startCol; c <= range.endCol; c++) {
        const sourceValues: string[] = [];
        for (let r = range.startRow; r <= range.endRow; r++) {
          sourceValues.push(this.state.getCellDisplay(c, r));
        }
        const series = detectSeries(sourceValues);
        for (let r = range.endRow + 1; r <= dragEndRow; r++) {
          const idx = r - range.startRow;
          const value = generateSeriesValue(series, sourceValues, idx);
          messageBridge.postMessage({
            type: 'editCell',
            sheet: sheet.name,
            col: c,
            row: r,
            value,
          });
        }
      }
    } else if (dragEndCol > range.endCol) {
      // Fill right
      for (let r = range.startRow; r <= range.endRow; r++) {
        const sourceValues: string[] = [];
        for (let c = range.startCol; c <= range.endCol; c++) {
          sourceValues.push(this.state.getCellDisplay(c, r));
        }
        const series = detectSeries(sourceValues);
        for (let c = range.endCol + 1; c <= dragEndCol; c++) {
          const idx = c - range.startCol;
          const value = generateSeriesValue(series, sourceValues, idx);
          messageBridge.postMessage({
            type: 'editCell',
            sheet: sheet.name,
            col: c,
            row: r,
            value,
          });
        }
      }
    } else if (dragEndRow < range.startRow) {
      // Fill up
      for (let c = range.startCol; c <= range.endCol; c++) {
        const sourceValues: string[] = [];
        for (let r = range.endRow; r >= range.startRow; r--) {
          sourceValues.push(this.state.getCellDisplay(c, r));
        }
        const series = detectSeries(sourceValues);
        for (let r = range.startRow - 1; r >= dragEndRow; r--) {
          const idx = range.startRow - r;
          const fillIdx = range.endRow - range.startRow + idx;
          const value = generateSeriesValue(series, sourceValues, fillIdx);
          messageBridge.postMessage({
            type: 'editCell',
            sheet: sheet.name,
            col: c,
            row: r,
            value,
          });
        }
      }
    }
  }

  private handleDblClick(e: MouseEvent): void {
    const { x, y } = this.getCanvasCoords(e);

    // Double-click on column resize handle → auto-fit
    const resizeCol = this.renderer.hitTestColResize(x, y);
    if (resizeCol !== null) {
      this.autoFitColumn(resizeCol);
      return;
    }

    const hit = this.renderer.hitTest(x, y);
    if (hit) {
      this.selection.selectCell(hit.col, hit.row);
      this.editor.startEditing();
      this.renderer.markDirty();
    }
  }

  private autoFitColumn(col: number): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;

    let maxWidth = 40; // minimum width
    // Scan visible/used rows for this column
    const cellLayer = this.renderer.cellLayer;
    for (const [key, cell] of sheet.cells) {
      const idx = key.indexOf(',');
      const c = parseInt(key.substring(0, idx), 10);
      if (c !== col) continue;
      const r = parseInt(key.substring(idx + 1), 10);

      const val = cell.computedValue ?? cell.rawValue;
      if (val === null) continue;
      const text = String(val);
      if (!text) continue;

      const style = this.state.getCellStyle(c, r);
      const w = cellLayer.measureCellTextWidth(text, style);
      if (w > maxWidth) maxWidth = w;
    }

    maxWidth = Math.min(400, Math.ceil(maxWidth + 4)); // cap + small padding
    sheet.columnWidths[col] = maxWidth;
    this.renderer.updateViewport();
    this.renderer.markDirty();

    messageBridge.postMessage({
      type: 'resizeColumn',
      sheet: sheet.name,
      col,
      width: maxWidth,
    });
  }

  private handleKeyDown(e: KeyboardEvent): void {
    // Don't handle keys while editing (CellEditor handles those)
    if (this.state.isEditing) return;

    // Don't handle keys when focus is on input elements
    const tag = (e.target as HTMLElement)?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;

    const shift = e.shiftKey;
    const ctrl = e.ctrlKey || e.metaKey;

    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        this.selection.moveSelection(0, -1, shift);
        this.scrollToCell(this.state.selectedCol, this.state.selectionEndRow);
        this.renderer.markDirty();
        break;
      case 'ArrowDown':
        e.preventDefault();
        this.selection.moveSelection(0, 1, shift);
        this.scrollToCell(this.state.selectedCol, this.state.selectionEndRow);
        this.renderer.markDirty();
        break;
      case 'ArrowLeft':
        e.preventDefault();
        this.selection.moveSelection(-1, 0, shift);
        this.scrollToCell(this.state.selectionEndCol, this.state.selectedRow);
        this.renderer.markDirty();
        break;
      case 'ArrowRight':
        e.preventDefault();
        this.selection.moveSelection(1, 0, shift);
        this.scrollToCell(this.state.selectionEndCol, this.state.selectedRow);
        this.renderer.markDirty();
        break;
      case 'Home':
        e.preventDefault();
        if (ctrl) {
          this.selection.selectCell(0, 0);
        } else {
          this.selection.selectCell(0, this.state.selectedRow);
        }
        this.scrollToCell(0, this.state.selectedRow);
        this.renderer.markDirty();
        break;
      case 'End': {
        e.preventDefault();
        const sheet = this.state.activeSheet;
        if (ctrl && sheet) {
          // Find last used row/col
          let maxC = 0;
          let maxR = 0;
          for (const key of sheet.cells.keys()) {
            const idx = key.indexOf(',');
            const cc = parseInt(key.substring(0, idx), 10);
            const rr = parseInt(key.substring(idx + 1), 10);
            if (cc > maxC) maxC = cc;
            if (rr > maxR) maxR = rr;
          }
          this.selection.selectCell(maxC, maxR);
          this.scrollToCell(maxC, maxR);
        } else {
          // End of current row — go to last used cell in this row
          const row = this.state.selectedRow;
          let maxCol = this.state.selectedCol;
          if (sheet) {
            for (const key of sheet.cells.keys()) {
              const idx = key.indexOf(',');
              const cc = parseInt(key.substring(0, idx), 10);
              const rr = parseInt(key.substring(idx + 1), 10);
              if (rr === row && cc > maxCol) maxCol = cc;
            }
          }
          this.selection.selectCell(maxCol, row);
          this.scrollToCell(maxCol, row);
        }
        this.renderer.markDirty();
        break;
      }
      case 'PageDown': {
        e.preventDefault();
        const pageRows = Math.max(1, Math.floor(this.renderer.viewportHeight / 24) - 1);
        this.selection.moveSelection(0, pageRows, shift);
        this.scrollToCell(this.state.selectedCol, this.state.selectionEndRow);
        this.renderer.markDirty();
        break;
      }
      case 'PageUp': {
        e.preventDefault();
        const pageRowsUp = Math.max(1, Math.floor(this.renderer.viewportHeight / 24) - 1);
        this.selection.moveSelection(0, -pageRowsUp, shift);
        this.scrollToCell(this.state.selectedCol, this.state.selectionEndRow);
        this.renderer.markDirty();
        break;
      }
      case 'Tab':
        e.preventDefault();
        this.selection.moveSelection(shift ? -1 : 1, 0);
        this.scrollToCell(this.state.selectedCol, this.state.selectedRow);
        this.renderer.markDirty();
        break;
      case 'Enter':
        e.preventDefault();
        if (shift) {
          this.selection.moveSelection(0, -1);
        } else {
          this.editor.startEditing();
        }
        this.scrollToCell(this.state.selectedCol, this.state.selectedRow);
        this.renderer.markDirty();
        break;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        this.deleteCellContent();
        break;
      case 'F2':
        e.preventDefault();
        this.editor.startEditing();
        break;
      case 'd':
        if (ctrl) {
          e.preventDefault();
          this.fillDown();
        } else {
          this.startTyping(e.key);
        }
        break;
      case 'r':
        if (ctrl) {
          e.preventDefault();
          this.fillRight();
        } else {
          this.startTyping(e.key);
        }
        break;
      case 'a':
        if (ctrl) {
          e.preventDefault();
          this.selection.selectAll();
          this.renderer.markDirty();
        } else {
          this.startTyping(e.key);
        }
        break;
      case 'c':
        if (!ctrl) this.startTyping(e.key);
        // Ctrl+C handled by DOM copy event
        break;
      case 'v':
        if (ctrl && shift) {
          e.preventDefault();
          this.pasteValuesOnly();
        } else if (!ctrl) {
          this.startTyping(e.key);
        }
        // plain Ctrl+V handled by DOM paste event
        break;
      case 'x':
        if (!ctrl) this.startTyping(e.key);
        // Ctrl+X handled by DOM cut event
        break;
      case 'z':
        // Let VSCode handle undo/redo
        break;
      default:
        if (e.key.length === 1 && !ctrl) {
          this.startTyping(e.key);
        }
        break;
    }
  }

  /** Ctrl+D — copy top row of selection down to all other rows */
  private fillDown(): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range || range.startRow === range.endRow) return;

    for (let c = range.startCol; c <= range.endCol; c++) {
      const sourceValue = this.getCellClipboardValue(c, range.startRow);
      for (let r = range.startRow + 1; r <= range.endRow; r++) {
        messageBridge.postMessage({
          type: 'editCell',
          sheet: sheet.name,
          col: c,
          row: r,
          value: sourceValue,
        });
      }
    }
  }

  /** Ctrl+R — copy left column of selection right to all other columns */
  private fillRight(): void {
    const sheet = this.state.activeSheet;
    const range = this.state.selectionRange;
    if (!sheet || !range || range.startCol === range.endCol) return;

    for (let r = range.startRow; r <= range.endRow; r++) {
      const sourceValue = this.getCellClipboardValue(range.startCol, r);
      for (let c = range.startCol + 1; c <= range.endCol; c++) {
        messageBridge.postMessage({
          type: 'editCell',
          sheet: sheet.name,
          col: c,
          row: r,
          value: sourceValue,
        });
      }
    }
  }

  private startTyping(char: string): void {
    this.editor.startEditing(char);
  }

  private deleteCellContent(): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    const range = this.state.selectionRange;
    if (!range) return;

    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        messageBridge.postMessage({
          type: 'editCell',
          sheet: sheet.name,
          col: c,
          row: r,
          value: '',
        });
      }
    }
  }

  // ── Clipboard via DOM events ──────────────────────────────────────────────
  // Internal clipboard stores raw formulas/values and style ids
  private internalClipboard: Array<{
    col: number;
    row: number;
    value: string;
    styleId: string | null;
  }> = [];

  /** Collect clipboard data: formula if available, else raw value */
  private getCellClipboardValue(col: number, row: number): string {
    const cell = this.state.getCell(col, row);
    if (cell.formula) return '=' + cell.formula;
    const v = cell.rawValue;
    return v === null ? '' : String(v);
  }

  private buildClipboardText(): string {
    const range = this.state.selectionRange;
    if (!range) return '';
    const lines: string[] = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row: string[] = [];
      for (let c = range.startCol; c <= range.endCol; c++) {
        row.push(this.state.getCellDisplay(c, r)); // tab-text for external
      }
      lines.push(row.join('\t'));
    }
    return lines.join('\n');
  }

  private handleCopy(e: ClipboardEvent): void {
    if (this.state.isEditing) return;
    const tag = (document.activeElement as HTMLElement)?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;

    const range = this.state.selectionRange;
    const sheet = this.state.activeSheet;
    if (!range || !sheet) return;

    e.preventDefault();

    // Store formulas + styles internally
    this.internalClipboard = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const cell = this.state.getCell(c, r);
        this.internalClipboard.push({
          col: c - range.startCol,
          row: r - range.startRow,
          value: this.getCellClipboardValue(c, r),
          styleId: cell.styleId,
        });
      }
    }

    // Write display text to system clipboard for external apps
    e.clipboardData?.setData('text/plain', this.buildClipboardText());
  }

  private handleCut(e: ClipboardEvent): void {
    if (this.state.isEditing) return;
    const tag = (document.activeElement as HTMLElement)?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;

    this.handleCopy(e);
    this.deleteCellContent();
  }

  private handlePaste(e: ClipboardEvent): void {
    if (this.state.isEditing) return;
    const tag = (document.activeElement as HTMLElement)?.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA') return;

    const sheet = this.state.activeSheet;
    if (!sheet) return;

    e.preventDefault();

    const text = e.clipboardData?.getData('text/plain') ?? '';

    if (text) {
      // Paste from system clipboard (external or previously copied)
      const rows = text.split(/\r?\n/);
      // Remove trailing empty row (common in Excel copy)
      if (rows.length > 0 && rows[rows.length - 1] === '') rows.pop();

      const startCol = this.state.selectedCol;
      const startRow = this.state.selectedRow;
      for (let r = 0; r < rows.length; r++) {
        const cols = rows[r].split('\t');
        for (let c = 0; c < cols.length; c++) {
          messageBridge.postMessage({
            type: 'editCell',
            sheet: sheet.name,
            col: startCol + c,
            row: startRow + r,
            value: cols[c],
          });
        }
      }
    } else if (this.internalClipboard.length > 0) {
      // Fallback to internal clipboard (with formulas)
      for (const item of this.internalClipboard) {
        messageBridge.postMessage({
          type: 'editCell',
          sheet: sheet.name,
          col: this.state.selectedCol + item.col,
          row: this.state.selectedRow + item.row,
          value: item.value,
        });
      }
    }
  }

  /** Paste values only (strips formulas and styles) */
  pasteValuesOnly(): void {
    const sheet = this.state.activeSheet;
    if (!sheet || this.internalClipboard.length === 0) return;
    messageBridge.postMessage({
      type: 'pasteSpecial',
      sheet: sheet.name,
      startCol: this.state.selectedCol,
      startRow: this.state.selectedRow,
      cells: this.internalClipboard.map((item) => ({
        col: item.col,
        row: item.row,
        value: item.value,
        styleId: null,
      })),
      mode: 'valuesOnly',
    });
  }

  hasClipboard(): boolean {
    return this.internalClipboard.length > 0;
  }

  /** Paste formats only (styles, no content) */
  pasteFormatsOnly(): void {
    const sheet = this.state.activeSheet;
    if (!sheet || this.internalClipboard.length === 0) return;
    messageBridge.postMessage({
      type: 'pasteSpecial',
      sheet: sheet.name,
      startCol: this.state.selectedCol,
      startRow: this.state.selectedRow,
      cells: this.internalClipboard.map((item) => ({
        col: item.col,
        row: item.row,
        value: '',
        styleId: item.styleId ?? null,
      })),
      mode: 'formatsOnly',
    });
  }
}

// ── Series detection and generation ────────────────────────────────────────

interface SeriesInfo {
  type: 'number' | 'ip' | 'letter' | 'repeat';
  step?: number;
  baseValues?: string[];
}

function detectSeries(values: string[]): SeriesInfo {
  if (values.length === 0) return { type: 'repeat' };

  // Try numeric series
  const nums = values.map(Number);
  if (nums.every((n) => !isNaN(n)) && values.every((v) => v !== '')) {
    if (values.length >= 2) {
      const step = nums[1] - nums[0];
      const isArithmetic = nums.every(
        (n, i) => i === 0 || Math.abs(n - nums[i - 1] - step) < 1e-10,
      );
      if (isArithmetic) return { type: 'number', step };
    }
    return { type: 'number', step: 1 };
  }

  // Try IP address series
  const ipRegex = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/;
  if (values.every((v) => ipRegex.test(v))) {
    return { type: 'ip' };
  }

  // Try single letter series (a, b, c or A, B, C)
  if (values.every((v) => /^[a-zA-Z]$/.test(v)) && values.length >= 2) {
    const codes = values.map((v) => v.charCodeAt(0));
    const step = codes[1] - codes[0];
    const isSequence = codes.every((c, i) => i === 0 || c - codes[i - 1] === step);
    if (isSequence) return { type: 'letter', step };
  }

  // Default: repeat pattern
  return { type: 'repeat', baseValues: values };
}

function generateSeriesValue(series: SeriesInfo, sourceValues: string[], index: number): string {
  switch (series.type) {
    case 'number': {
      const start = Number(sourceValues[0]);
      const step = series.step ?? 1;
      const result = start + step * index;
      // Preserve decimal precision from source
      const decimalPlaces = Math.max(
        ...sourceValues.map((v) => {
          const d = v.split('.');
          return d.length > 1 ? d[1].length : 0;
        }),
      );
      return decimalPlaces > 0 ? result.toFixed(decimalPlaces) : String(result);
    }

    case 'ip': {
      const parts = sourceValues[0].split('.').map(Number);
      if (sourceValues.length >= 2) {
        const parts2 = sourceValues[1].split('.').map(Number);
        // Find which octet changes
        for (let o = 3; o >= 0; o--) {
          if (parts2[o] !== parts[o]) {
            const step = parts2[o] - parts[o];
            const result = [...parts];
            result[o] = parts[o] + step * index;
            // Handle overflow
            if (result[o] > 255) result[o] = result[o] % 256;
            if (result[o] < 0) result[o] = 256 + (result[o] % 256);
            return result.join('.');
          }
        }
      }
      // Default: increment last octet
      const result = [...parts];
      result[3] = (parts[3] + index) % 256;
      return result.join('.');
    }

    case 'letter': {
      const startCode = sourceValues[0].charCodeAt(0);
      const step = series.step ?? 1;
      const code = startCode + step * index;
      // Wrap around a-z or A-Z
      if (sourceValues[0] >= 'a' && sourceValues[0] <= 'z') {
        return String.fromCharCode(((code - 97 + 26) % 26) + 97);
      }
      return String.fromCharCode(((code - 65 + 26) % 26) + 65);
    }

    case 'repeat':
    default: {
      const base = series.baseValues ?? sourceValues;
      return base[index % base.length];
    }
  }
}
