import { ViewportCalculator } from './ViewportCalculator';

const SCROLL_SPEED = 1;

export class ScrollManager {
  scrollX = 0;
  scrollY = 0;

  private onScroll: () => void;
  private viewport: ViewportCalculator;
  private canvasWidth = 0;
  private canvasHeight = 0;

  constructor(canvas: HTMLCanvasElement, viewport: ViewportCalculator, onScroll: () => void) {
    this.viewport = viewport;
    this.onScroll = onScroll;

    canvas.addEventListener('wheel', (e) => this.handleWheel(e), { passive: false });
  }

  setCanvasSize(width: number, height: number): void {
    this.canvasWidth = width;
    this.canvasHeight = height;
  }

  private handleWheel(e: WheelEvent): void {
    e.preventDefault();

    const dx = e.deltaX * SCROLL_SPEED;
    const dy = e.deltaY * SCROLL_SPEED;

    const maxScrollX = Math.max(0, this.viewport.totalWidth - (this.canvasWidth - 50));
    const maxScrollY = Math.max(0, this.viewport.totalHeight - (this.canvasHeight - 24));

    this.scrollX = Math.max(0, Math.min(maxScrollX, this.scrollX + dx));
    this.scrollY = Math.max(0, Math.min(maxScrollY, this.scrollY + dy));

    this.onScroll();
  }

  /** Ensure a cell is visible, scrolling if necessary.
   *  Frozen cells are always visible — don't adjust scroll for them. */
  scrollToCell(col: number, row: number, frozenCols = 0, frozenRows = 0): void {
    const frozenColWidth = frozenCols > 0 ? this.viewport.colLeft(frozenCols) : 0;
    const frozenRowHeight = frozenRows > 0 ? this.viewport.rowTop(frozenRows) : 0;

    const viewW = this.canvasWidth - 50 - frozenColWidth;
    const viewH = this.canvasHeight - 24 - frozenRowHeight;

    let changed = false;

    // Only scroll horizontally for non-frozen columns
    if (col >= frozenCols) {
      const cellLeft = this.viewport.colLeft(col) - frozenColWidth;
      const cellRight = cellLeft + this.viewport.colWidth(col);
      if (cellLeft < this.scrollX) {
        this.scrollX = cellLeft;
        changed = true;
      } else if (cellRight > this.scrollX + viewW) {
        this.scrollX = cellRight - viewW;
        changed = true;
      }
    }

    // Only scroll vertically for non-frozen rows
    if (row >= frozenRows) {
      const cellTop = this.viewport.rowTop(row) - frozenRowHeight;
      const cellBottom = cellTop + this.viewport.rowHeight(row);
      if (cellTop < this.scrollY) {
        this.scrollY = cellTop;
        changed = true;
      } else if (cellBottom > this.scrollY + viewH) {
        this.scrollY = cellBottom - viewH;
        changed = true;
      }
    }

    if (changed) this.onScroll();
  }
}
