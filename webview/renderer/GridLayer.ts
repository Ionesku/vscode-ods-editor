import {
  ViewportCalculator,
  VisibleRange,
  ROW_HEADER_WIDTH,
  COL_HEADER_HEIGHT,
} from './ViewportCalculator';
import { colToLetter } from '../../shared/types';

export class GridLayer {
  private headerBg = '#f8f9fa';
  private headerText = '#5f6368';
  private gridLine = '#e2e2e2';
  private headerBorder = '#c0c0c0';

  draw(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    range: VisibleRange,
    scrollX: number,
    scrollY: number,
    canvasWidth: number,
    canvasHeight: number,
    frozenRows = 0,
    frozenCols = 0,
  ): void {
    const frozenY =
      frozenRows > 0 ? COL_HEADER_HEIGHT + viewport.rowTop(frozenRows) : COL_HEADER_HEIGHT;
    const frozenX =
      frozenCols > 0 ? ROW_HEADER_WIDTH + viewport.colLeft(frozenCols) : ROW_HEADER_WIDTH;

    // Draw grid lines for scrollable area (clipped to non-frozen zone)
    ctx.save();
    ctx.beginPath();
    ctx.rect(frozenX, frozenY, canvasWidth - frozenX, canvasHeight - frozenY);
    ctx.clip();
    ctx.strokeStyle = this.gridLine;
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let c = range.startCol; c <= range.endCol; c++) {
      if (c < frozenCols) continue;
      const x = ROW_HEADER_WIDTH + viewport.colLeft(c) - scrollX + viewport.colWidth(c);
      ctx.moveTo(Math.round(x) + 0.5, frozenY);
      ctx.lineTo(Math.round(x) + 0.5, canvasHeight);
    }
    for (let r = range.startRow; r <= range.endRow; r++) {
      if (r < frozenRows) continue;
      const y = COL_HEADER_HEIGHT + viewport.rowTop(r) - scrollY + viewport.rowHeight(r);
      ctx.moveTo(frozenX, Math.round(y) + 0.5);
      ctx.lineTo(canvasWidth, Math.round(y) + 0.5);
    }
    ctx.stroke();
    ctx.restore();

    // Draw frozen grid lines (they don't scroll)
    if (frozenRows > 0 || frozenCols > 0) {
      ctx.strokeStyle = this.gridLine;
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (let c = 0; c < frozenCols; c++) {
        const x = ROW_HEADER_WIDTH + viewport.colLeft(c) + viewport.colWidth(c);
        ctx.moveTo(Math.round(x) + 0.5, COL_HEADER_HEIGHT);
        ctx.lineTo(Math.round(x) + 0.5, canvasHeight);
      }
      for (let r = 0; r < frozenRows; r++) {
        const y = COL_HEADER_HEIGHT + viewport.rowTop(r) + viewport.rowHeight(r);
        ctx.moveTo(ROW_HEADER_WIDTH, Math.round(y) + 0.5);
        ctx.lineTo(canvasWidth, Math.round(y) + 0.5);
      }
      ctx.stroke();
    }
  }

  /** Draw freeze pane divider lines (on top of everything) */
  drawFreezeLines(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    canvasWidth: number,
    canvasHeight: number,
    frozenRows: number,
    frozenCols: number,
  ): void {
    ctx.strokeStyle = '#888';
    ctx.lineWidth = 2;
    if (frozenCols > 0) {
      const fx = ROW_HEADER_WIDTH + viewport.colLeft(frozenCols);
      ctx.beginPath();
      ctx.moveTo(fx, 0);
      ctx.lineTo(fx, canvasHeight);
      ctx.stroke();
    }
    if (frozenRows > 0) {
      const fy = COL_HEADER_HEIGHT + viewport.rowTop(frozenRows);
      ctx.beginPath();
      ctx.moveTo(0, fy);
      ctx.lineTo(canvasWidth, fy);
      ctx.stroke();
    }
  }

  /** Re-draw headers on top of cells so frozen cells don't bleed into header area */
  drawHeaders(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    range: VisibleRange,
    scrollX: number,
    scrollY: number,
    canvasWidth: number,
    canvasHeight: number,
    frozenRows: number,
    frozenCols: number,
  ): void {
    // Row header background
    ctx.fillStyle = this.headerBg;
    ctx.fillRect(0, 0, ROW_HEADER_WIDTH, canvasHeight);
    // Col header background
    ctx.fillRect(0, 0, canvasWidth, COL_HEADER_HEIGHT);

    ctx.fillStyle = this.headerText;
    ctx.font = '11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    const frozenY =
      frozenRows > 0 ? COL_HEADER_HEIGHT + viewport.rowTop(frozenRows) : COL_HEADER_HEIGHT;
    const frozenX =
      frozenCols > 0 ? ROW_HEADER_WIDTH + viewport.colLeft(frozenCols) : ROW_HEADER_WIDTH;

    // Scrollable col headers (clipped to right of frozen cols)
    ctx.save();
    ctx.beginPath();
    ctx.rect(frozenX, 0, canvasWidth - frozenX, COL_HEADER_HEIGHT);
    ctx.clip();
    for (let c = range.startCol; c <= range.endCol; c++) {
      if (c < frozenCols) continue;
      const x = ROW_HEADER_WIDTH + viewport.colLeft(c) - scrollX;
      const w = viewport.colWidth(c);
      ctx.fillText(colToLetter(c), x + w / 2, COL_HEADER_HEIGHT / 2, w);
    }
    ctx.restore();

    // Frozen col headers (on top, no scroll)
    ctx.fillStyle = this.headerText;
    ctx.font = '11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    for (let c = 0; c < frozenCols; c++) {
      const x = ROW_HEADER_WIDTH + viewport.colLeft(c);
      const w = viewport.colWidth(c);
      ctx.fillText(colToLetter(c), x + w / 2, COL_HEADER_HEIGHT / 2, w);
    }

    // Scrollable row headers (clipped to below frozen rows)
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, frozenY, ROW_HEADER_WIDTH, canvasHeight - frozenY);
    ctx.clip();
    ctx.fillStyle = this.headerText;
    ctx.font = '11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    for (let r = range.startRow; r <= range.endRow; r++) {
      if (r < frozenRows) continue;
      const y = COL_HEADER_HEIGHT + viewport.rowTop(r) - scrollY;
      const h = viewport.rowHeight(r);
      ctx.fillText(String(r + 1), ROW_HEADER_WIDTH / 2, y + h / 2, ROW_HEADER_WIDTH - 4);
    }
    ctx.restore();

    // Frozen row headers (on top, no scroll)
    ctx.fillStyle = this.headerText;
    ctx.font = '11px -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    for (let r = 0; r < frozenRows; r++) {
      const y = COL_HEADER_HEIGHT + viewport.rowTop(r);
      const h = viewport.rowHeight(r);
      ctx.fillText(String(r + 1), ROW_HEADER_WIDTH / 2, y + h / 2, ROW_HEADER_WIDTH - 4);
    }

    // Header borders
    ctx.strokeStyle = this.headerBorder;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(0, COL_HEADER_HEIGHT + 0.5);
    ctx.lineTo(canvasWidth, COL_HEADER_HEIGHT + 0.5);
    ctx.moveTo(ROW_HEADER_WIDTH + 0.5, 0);
    ctx.lineTo(ROW_HEADER_WIDTH + 0.5, canvasHeight);
    ctx.stroke();

    // Top-left corner
    ctx.fillStyle = this.headerBg;
    ctx.fillRect(0, 0, ROW_HEADER_WIDTH, COL_HEADER_HEIGHT);
    ctx.strokeStyle = this.headerBorder;
    ctx.strokeRect(0.5, 0.5, ROW_HEADER_WIDTH, COL_HEADER_HEIGHT);
  }

  setTheme(colors: {
    headerBg: string;
    headerText: string;
    gridLine: string;
    headerBorder: string;
  }): void {
    this.headerBg = colors.headerBg;
    this.headerText = colors.headerText;
    this.gridLine = colors.gridLine;
    this.headerBorder = colors.headerBorder;
  }
}
