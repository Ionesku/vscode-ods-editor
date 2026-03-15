import { ViewportCalculator, ROW_HEADER_WIDTH, COL_HEADER_HEIGHT } from './ViewportCalculator';
import { WebviewState } from '../state/WebviewState';

export class SelectionLayer {
  private selectionColor = 'rgba(38, 79, 120, 0.3)';
  private selectionBorder = '#264f78';
  private activeCellBorder = '#007acc';

  draw(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    scrollX: number,
    scrollY: number,
    state: WebviewState,
    frozenRows = 0,
    frozenCols = 0,
  ): void {
    const range = state.selectionRange;
    if (!range) return;

    // Frozen cells don't scroll; non-frozen cells do
    const cellScreenX = (col: number) =>
      col < frozenCols
        ? ROW_HEADER_WIDTH + viewport.colLeft(col)
        : ROW_HEADER_WIDTH + viewport.colLeft(col) - scrollX;
    const cellScreenY = (row: number) =>
      row < frozenRows
        ? COL_HEADER_HEIGHT + viewport.rowTop(row)
        : COL_HEADER_HEIGHT + viewport.rowTop(row) - scrollY;

    const x1 = cellScreenX(range.startCol);
    const y1 = cellScreenY(range.startRow);
    const x2 = cellScreenX(range.endCol) + viewport.colWidth(range.endCol);
    const y2 = cellScreenY(range.endRow) + viewport.rowHeight(range.endRow);
    const w = x2 - x1;
    const h = y2 - y1;

    // Draw selection fill
    ctx.fillStyle = this.selectionColor;
    ctx.fillRect(x1, y1, w, h);

    // Draw selection border
    ctx.strokeStyle = this.selectionBorder;
    ctx.lineWidth = 2;
    ctx.strokeRect(x1, y1, w, h);

    // Draw active cell (thicker border) — span the full merge area if merged
    const activeX = cellScreenX(state.selectedCol);
    const activeY = cellScreenY(state.selectedRow);
    const activeCell = state.getCell(state.selectedCol, state.selectedRow);
    const mergeColSpan = activeCell?.mergeColSpan ?? 1;
    const mergeRowSpan = activeCell?.mergeRowSpan ?? 1;
    let activeW = 0;
    for (let c = state.selectedCol; c < state.selectedCol + mergeColSpan; c++) {
      activeW += viewport.colWidth(c);
    }
    let activeH = 0;
    for (let r = state.selectedRow; r < state.selectedRow + mergeRowSpan; r++) {
      activeH += viewport.rowHeight(r);
    }

    ctx.strokeStyle = this.activeCellBorder;
    ctx.lineWidth = 2;
    ctx.strokeRect(activeX, activeY, activeW, activeH);

    // Fill handle (small square at bottom-right of selection)
    const handleSize = 6;
    ctx.fillStyle = this.activeCellBorder;
    ctx.fillRect(x1 + w - handleSize / 2, y1 + h - handleSize / 2, handleSize, handleSize);

    // Highlight row and column headers
    this.highlightHeaders(ctx, viewport, scrollX, scrollY, state, frozenRows, frozenCols);
  }

  private highlightHeaders(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    scrollX: number,
    scrollY: number,
    state: WebviewState,
    frozenRows: number,
    frozenCols: number,
  ): void {
    const range = state.selectionRange;
    if (!range) return;

    ctx.fillStyle = 'rgba(38, 79, 120, 0.4)';

    for (let c = range.startCol; c <= range.endCol; c++) {
      const x =
        c < frozenCols
          ? ROW_HEADER_WIDTH + viewport.colLeft(c)
          : ROW_HEADER_WIDTH + viewport.colLeft(c) - scrollX;
      const w = viewport.colWidth(c);
      ctx.fillRect(x, 0, w, COL_HEADER_HEIGHT);
    }

    for (let r = range.startRow; r <= range.endRow; r++) {
      const y =
        r < frozenRows
          ? COL_HEADER_HEIGHT + viewport.rowTop(r)
          : COL_HEADER_HEIGHT + viewport.rowTop(r) - scrollY;
      const h = viewport.rowHeight(r);
      ctx.fillRect(0, y, ROW_HEADER_WIDTH, h);
    }
  }
}
