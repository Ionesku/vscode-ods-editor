import { GridLayer } from './GridLayer';
import { CellLayer } from './CellLayer';
import { SelectionLayer } from './SelectionLayer';
import { ViewportCalculator, ROW_HEADER_WIDTH, COL_HEADER_HEIGHT } from './ViewportCalculator';
import { ScrollManager } from './ScrollManager';
import { WebviewState } from '../state/WebviewState';

export class CanvasRenderer {
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private gridLayer: GridLayer;
  cellLayer: CellLayer;
  private selectionLayer: SelectionLayer;
  viewport: ViewportCalculator;
  scrollManager: ScrollManager;
  private rafId: number | null = null;
  private selectionRafId: number | null = null;
  private dpr: number;

  /** Offscreen canvas caching the grid+cells layer (everything except selection) */
  private offscreenCanvas: OffscreenCanvas | null = null;
  private offscreenCtx: OffscreenCanvasRenderingContext2D | null = null;
  private offscreenValid = false;

  constructor(
    canvas: HTMLCanvasElement,
    private state: WebviewState,
  ) {
    this.canvas = canvas;
    const ctx = canvas.getContext('2d');
    if (!ctx) throw new Error('Could not get 2D rendering context from canvas');
    this.ctx = ctx;
    this.dpr = window.devicePixelRatio || 1;

    this.gridLayer = new GridLayer();
    this.cellLayer = new CellLayer(this.ctx);
    this.selectionLayer = new SelectionLayer();
    this.viewport = new ViewportCalculator();
    this.scrollManager = new ScrollManager(canvas, this.viewport, () => this.markDirty());
  }

  /** Resize canvas to fill container */
  resize(): void {
    const container = this.canvas.parentElement;
    if (!container) return;
    const rect = container.getBoundingClientRect();
    const w = rect.width;
    const h = rect.height;

    this.dpr = window.devicePixelRatio || 1;
    this.canvas.width = w * this.dpr;
    this.canvas.height = h * this.dpr;
    this.canvas.style.width = w + 'px';
    this.canvas.style.height = h + 'px';

    // Recreate offscreen canvas to match new size
    this.offscreenCanvas = new OffscreenCanvas(w * this.dpr, h * this.dpr);
    this.offscreenCtx = this.offscreenCanvas.getContext('2d') as OffscreenCanvasRenderingContext2D;
    this.offscreenValid = false;

    this.scrollManager.setCanvasSize(w, h);
    this.markDirty();
  }

  get viewportHeight(): number {
    return this.canvas.clientHeight;
  }

  /** Update viewport with new sheet data */
  updateViewport(): void {
    const sheet = this.state.activeSheet;
    if (!sheet) return;
    this.viewport.updateColumnWidths(sheet.columnWidths);
    this.viewport.updateRowHeights(sheet.rowHeights);
  }

  /** Full redraw: grid + cells + selection */
  markDirty(): void {
    this.offscreenValid = false;
    if (this.rafId === null) {
      this.rafId = requestAnimationFrame(() => this.render());
    }
  }

  /**
   * Light redraw: only re-composite the selection layer on top of the cached
   * grid+cells image. Use this when only the selection has changed (e.g. arrow
   * key navigation) to avoid redrawing all cells every keystroke.
   */
  markSelectionDirty(): void {
    if (this.offscreenValid && this.selectionRafId === null) {
      this.selectionRafId = requestAnimationFrame(() => this.renderSelectionOnly());
    } else {
      // Fall back to full redraw if background isn't cached yet
      this.markDirty();
    }
  }

  private renderSelectionOnly(): void {
    this.selectionRafId = null;
    if (!this.offscreenValid || !this.offscreenCanvas) {
      this.render();
      return;
    }

    const ctx = this.ctx;
    const w = this.canvas.width / this.dpr;
    const h = this.canvas.height / this.dpr;

    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    // Blit cached background
    ctx.drawImage(this.offscreenCanvas, 0, 0, w, h, 0, 0, w, h);

    const sheet = this.state.activeSheet;
    if (!sheet) return;
    const scrollX = this.scrollManager.scrollX;
    const scrollY = this.scrollManager.scrollY;

    // Re-draw headers on top (so selection highlight in header shows correctly)
    const range = this.viewport.getVisibleRange(scrollX, scrollY, w, h);
    this.gridLayer.drawHeaders(ctx, this.viewport, range, scrollX, scrollY, w, h, sheet.frozenRows, sheet.frozenCols);

    this.selectionLayer.draw(ctx, this.viewport, scrollX, scrollY, this.state, sheet.frozenRows, sheet.frozenCols);
  }

  private render(): void {
    this.rafId = null;
    this.selectionRafId = null; // cancel any pending selection-only redraw
    const ctx = this.ctx;
    const w = this.canvas.width / this.dpr;
    const h = this.canvas.height / this.dpr;

    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);

    const sheet = this.state.activeSheet;
    if (!sheet) return;

    const scrollX = this.scrollManager.scrollX;
    const scrollY = this.scrollManager.scrollY;
    const frozenRows = sheet.frozenRows;
    const frozenCols = sheet.frozenCols;
    const range = this.viewport.getVisibleRange(scrollX, scrollY, w, h);

    // Layer 1: Grid (background, headers, gridlines)
    this.gridLayer.draw(ctx, this.viewport, range, scrollX, scrollY, w, h, frozenRows, frozenCols);

    // Layer 2: Cell content — scrollable area only (clipped to avoid frozen zones)
    if (frozenRows > 0 || frozenCols > 0) {
      const frozenX =
        frozenCols > 0 ? ROW_HEADER_WIDTH + this.viewport.colLeft(frozenCols) : ROW_HEADER_WIDTH;
      const frozenY =
        frozenRows > 0 ? COL_HEADER_HEIGHT + this.viewport.rowTop(frozenRows) : COL_HEADER_HEIGHT;

      // Draw scrollable cells clipped to non-frozen area
      ctx.save();
      ctx.beginPath();
      ctx.rect(frozenX, frozenY, w - frozenX, h - frozenY);
      ctx.clip();
      this.cellLayer.drawRegion(
        ctx,
        this.viewport,
        range,
        scrollX,
        scrollY,
        this.state,
        frozenCols,
        range.endCol,
        frozenRows,
        range.endRow,
      );
      ctx.restore();

      // Draw frozen-row + scrollable-col (top strip, clipped horizontally)
      if (frozenRows > 0) {
        ctx.save();
        // Solid background to cover scrollable content
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(frozenX, COL_HEADER_HEIGHT, w - frozenX, frozenY - COL_HEADER_HEIGHT);
        ctx.beginPath();
        ctx.rect(frozenX, COL_HEADER_HEIGHT, w - frozenX, frozenY - COL_HEADER_HEIGHT);
        ctx.clip();
        this.cellLayer.drawRegion(
          ctx,
          this.viewport,
          range,
          scrollX,
          0,
          this.state,
          frozenCols,
          range.endCol,
          0,
          frozenRows - 1,
        );
        ctx.restore();
      }

      // Draw frozen-col + scrollable-row (left strip, clipped vertically)
      if (frozenCols > 0) {
        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(ROW_HEADER_WIDTH, frozenY, frozenX - ROW_HEADER_WIDTH, h - frozenY);
        ctx.beginPath();
        ctx.rect(ROW_HEADER_WIDTH, frozenY, frozenX - ROW_HEADER_WIDTH, h - frozenY);
        ctx.clip();
        this.cellLayer.drawRegion(
          ctx,
          this.viewport,
          range,
          0,
          scrollY,
          this.state,
          0,
          frozenCols - 1,
          frozenRows,
          range.endRow,
        );
        ctx.restore();
      }

      // Draw frozen corner (top-left, fully frozen)
      if (frozenRows > 0 && frozenCols > 0) {
        ctx.save();
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(
          ROW_HEADER_WIDTH,
          COL_HEADER_HEIGHT,
          frozenX - ROW_HEADER_WIDTH,
          frozenY - COL_HEADER_HEIGHT,
        );
        ctx.beginPath();
        ctx.rect(
          ROW_HEADER_WIDTH,
          COL_HEADER_HEIGHT,
          frozenX - ROW_HEADER_WIDTH,
          frozenY - COL_HEADER_HEIGHT,
        );
        ctx.clip();
        this.cellLayer.drawRegion(
          ctx,
          this.viewport,
          range,
          0,
          0,
          this.state,
          0,
          frozenCols - 1,
          0,
          frozenRows - 1,
        );
        ctx.restore();
      }
    } else {
      this.cellLayer.draw(ctx, this.viewport, range, scrollX, scrollY, this.state, 0, 0);
    }

    // Re-draw grid frozen pane dividers on top of cells
    if (frozenCols > 0 || frozenRows > 0) {
      this.gridLayer.drawFreezeLines(ctx, this.viewport, w, h, frozenRows, frozenCols);
    }

    // Re-draw headers on top (so cells don't bleed into header area)
    this.gridLayer.drawHeaders(
      ctx,
      this.viewport,
      range,
      scrollX,
      scrollY,
      w,
      h,
      frozenRows,
      frozenCols,
    );

    // Cache background (grid + cells) before drawing selection
    if (this.offscreenCanvas && this.offscreenCtx) {
      this.offscreenCtx.setTransform(1, 0, 0, 1, 0, 0);
      this.offscreenCtx.drawImage(this.canvas, 0, 0);
      this.offscreenValid = true;
    }

    // Layer 3: Selection highlight
    this.selectionLayer.draw(
      ctx,
      this.viewport,
      scrollX,
      scrollY,
      this.state,
      frozenRows,
      frozenCols,
    );
  }

  /** Convert canvas pixel coordinates to cell col,row */
  hitTest(canvasX: number, canvasY: number): { col: number; row: number } | null {
    const x = canvasX - ROW_HEADER_WIDTH;
    const y = canvasY - COL_HEADER_HEIGHT;
    if (x < 0 || y < 0) return null;

    const sheet = this.state.activeSheet;
    const frozenCols = sheet?.frozenCols ?? 0;
    const frozenRows = sheet?.frozenRows ?? 0;

    // Check if click is in frozen column area
    const frozenColWidth = frozenCols > 0 ? this.viewport.colLeft(frozenCols) : 0;
    const frozenRowHeight = frozenRows > 0 ? this.viewport.rowTop(frozenRows) : 0;

    let col: number;
    if (frozenCols > 0 && x < frozenColWidth) {
      col = this.viewport.colAtX(x);
    } else {
      col = this.viewport.colAtX(
        x + this.scrollManager.scrollX - (frozenCols > 0 ? frozenColWidth - frozenColWidth : 0),
      );
    }

    let row: number;
    if (frozenRows > 0 && y < frozenRowHeight) {
      row = this.viewport.rowAtY(y);
    } else {
      row = this.viewport.rowAtY(
        y + this.scrollManager.scrollY - (frozenRows > 0 ? frozenRowHeight - frozenRowHeight : 0),
      );
    }

    return { col, row };
  }

  /** Check if click is on column header → returns column index */
  hitTestColHeader(canvasX: number, canvasY: number): number | null {
    if (canvasY >= COL_HEADER_HEIGHT || canvasX <= ROW_HEADER_WIDTH) return null;
    const sheet = this.state.activeSheet;
    const frozenCols = sheet?.frozenCols ?? 0;
    const x = canvasX - ROW_HEADER_WIDTH;
    const frozenColWidth = frozenCols > 0 ? this.viewport.colLeft(frozenCols) : 0;

    if (frozenCols > 0 && x < frozenColWidth) {
      return this.viewport.colAtX(x);
    }
    return this.viewport.colAtX(x + this.scrollManager.scrollX);
  }

  /** Check if click is on row header → returns row index */
  hitTestRowHeader(canvasX: number, canvasY: number): number | null {
    if (canvasX >= ROW_HEADER_WIDTH || canvasY <= COL_HEADER_HEIGHT) return null;
    const sheet = this.state.activeSheet;
    const frozenRows = sheet?.frozenRows ?? 0;
    const y = canvasY - COL_HEADER_HEIGHT;
    const frozenRowHeight = frozenRows > 0 ? this.viewport.rowTop(frozenRows) : 0;

    if (frozenRows > 0 && y < frozenRowHeight) {
      return this.viewport.rowAtY(y);
    }
    return this.viewport.rowAtY(y + this.scrollManager.scrollY);
  }

  /** Check if click is on column resize handle */
  hitTestColResize(canvasX: number, canvasY: number): number | null {
    if (canvasY > COL_HEADER_HEIGHT) return null;
    const HANDLE_WIDTH = 4;
    const range = this.viewport.getVisibleRange(
      this.scrollManager.scrollX,
      this.scrollManager.scrollY,
      this.canvas.width / this.dpr,
      this.canvas.height / this.dpr,
    );
    for (let c = range.startCol; c <= range.endCol; c++) {
      const rightEdge =
        ROW_HEADER_WIDTH +
        this.viewport.colLeft(c) -
        this.scrollManager.scrollX +
        this.viewport.colWidth(c);
      if (Math.abs(canvasX - rightEdge) < HANDLE_WIDTH) return c;
    }
    return null;
  }

  /** Check if click is on fill handle (bottom-right of selection) */
  hitTestFillHandle(canvasX: number, canvasY: number): boolean {
    const state = this.state;
    const range = state.selectionRange;
    if (!range) return false;

    const scrollX = this.scrollManager.scrollX;
    const scrollY = this.scrollManager.scrollY;

    let w = 0;
    for (let c = range.startCol; c <= range.endCol; c++) w += this.viewport.colWidth(c);
    let h = 0;
    for (let r = range.startRow; r <= range.endRow; r++) h += this.viewport.rowHeight(r);

    const x1 = ROW_HEADER_WIDTH + this.viewport.colLeft(range.startCol) - scrollX;
    const y1 = COL_HEADER_HEIGHT + this.viewport.rowTop(range.startRow) - scrollY;
    const handleX = x1 + w;
    const handleY = y1 + h;
    const tolerance = 6;

    return Math.abs(canvasX - handleX) <= tolerance && Math.abs(canvasY - handleY) <= tolerance;
  }

  /** Check if click is on row resize handle */
  hitTestRowResize(canvasX: number, canvasY: number): number | null {
    if (canvasX > ROW_HEADER_WIDTH) return null;
    const HANDLE_WIDTH = 4;
    const range = this.viewport.getVisibleRange(
      this.scrollManager.scrollX,
      this.scrollManager.scrollY,
      this.canvas.width / this.dpr,
      this.canvas.height / this.dpr,
    );
    for (let r = range.startRow; r <= range.endRow; r++) {
      const bottomEdge =
        COL_HEADER_HEIGHT +
        this.viewport.rowTop(r) -
        this.scrollManager.scrollY +
        this.viewport.rowHeight(r);
      if (Math.abs(canvasY - bottomEdge) < HANDLE_WIDTH) return r;
    }
    return null;
  }
}
