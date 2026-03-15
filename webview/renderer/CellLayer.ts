import {
  ViewportCalculator,
  VisibleRange,
  ROW_HEADER_WIDTH,
  COL_HEADER_HEIGHT,
} from './ViewportCalculator';
import { TextMeasurer } from './TextMeasurer';
import { WebviewState } from '../state/WebviewState';
import { CellStyle, BorderStyle } from '../../shared/types';

const CELL_PADDING = 4;

export class CellLayer {
  private textMeasurer: TextMeasurer;
  private defaultTextColor = '#202124';

  constructor(ctx: CanvasRenderingContext2D) {
    this.textMeasurer = new TextMeasurer(ctx);
  }

  draw(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    range: VisibleRange,
    scrollX: number,
    scrollY: number,
    state: WebviewState,
    frozenRows = 0,
    frozenCols = 0,
  ): void {
    // Draw scrollable cells (skip frozen area)
    for (let r = range.startRow; r <= range.endRow; r++) {
      if (r < frozenRows) continue;
      for (let c = range.startCol; c <= range.endCol; c++) {
        if (c < frozenCols) continue;
        const x = ROW_HEADER_WIDTH + viewport.colLeft(c) - scrollX;
        const y = COL_HEADER_HEIGHT + viewport.rowTop(r) - scrollY;
        this.drawCell(ctx, viewport, state, c, r, x, y);
      }
    }

    // Draw frozen-row + scrollable-col cells
    for (let r = 0; r < frozenRows; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        if (c < frozenCols) continue;
        const x = ROW_HEADER_WIDTH + viewport.colLeft(c) - scrollX;
        const y = COL_HEADER_HEIGHT + viewport.rowTop(r); // no scrollY
        this.drawCell(ctx, viewport, state, c, r, x, y);
      }
    }

    // Draw frozen-col + scrollable-row cells
    for (let r = range.startRow; r <= range.endRow; r++) {
      if (r < frozenRows) continue;
      for (let c = 0; c < frozenCols; c++) {
        const x = ROW_HEADER_WIDTH + viewport.colLeft(c); // no scrollX
        const y = COL_HEADER_HEIGHT + viewport.rowTop(r) - scrollY;
        this.drawCell(ctx, viewport, state, c, r, x, y);
      }
    }

    // Draw frozen-row + frozen-col cells (top-left corner)
    for (let r = 0; r < frozenRows; r++) {
      for (let c = 0; c < frozenCols; c++) {
        const x = ROW_HEADER_WIDTH + viewport.colLeft(c);
        const y = COL_HEADER_HEIGHT + viewport.rowTop(r);
        this.drawCell(ctx, viewport, state, c, r, x, y);
      }
    }
  }

  /** Draw cells in a specific col/row range with given scroll offsets */
  drawRegion(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    _visRange: VisibleRange,
    scrollX: number,
    scrollY: number,
    state: WebviewState,
    startCol: number,
    endCol: number,
    startRow: number,
    endRow: number,
  ): void {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        const x = ROW_HEADER_WIDTH + viewport.colLeft(c) - scrollX;
        const y = COL_HEADER_HEIGHT + viewport.rowTop(r) - scrollY;
        this.drawCell(ctx, viewport, state, c, r, x, y);
      }
    }
  }

  private drawCell(
    ctx: CanvasRenderingContext2D,
    viewport: ViewportCalculator,
    state: WebviewState,
    c: number,
    r: number,
    x: number,
    y: number,
  ): void {
    const sheet = state.activeSheet;
    if (sheet?.hiddenRows.has(r)) return;
    const cell = state.getCell(c, r);
    if (cell.mergedInto) return;

    let w = 0;
    for (let mc = 0; mc < cell.mergeColSpan; mc++) w += viewport.colWidth(c + mc);
    let h = 0;
    for (let mr = 0; mr < cell.mergeRowSpan; mr++) h += viewport.rowHeight(r + mr);

    const style = state.getCellStyle(c, r);

    if (cell.mergeColSpan > 1 || cell.mergeRowSpan > 1) {
      ctx.fillStyle = style?.backgroundColor ?? '#ffffff';
      ctx.fillRect(x, y, w, h);
    } else if (style?.backgroundColor) {
      ctx.fillStyle = style.backgroundColor;
      ctx.fillRect(x, y, w, h);
    }

    if (style) this.drawBorders(ctx, x, y, w, h, style);

    const display = state.getCellDisplay(c, r);
    if (display) this.drawCellText(ctx, x, y, w, h, display, cell.rawValue, style);
  }

  /** Measure text width for auto-fit */
  measureCellTextWidth(text: string, style?: CellStyle): number {
    const parts: string[] = [];
    if (style?.italic) parts.push('italic');
    if (style?.bold) parts.push('bold');
    const fontSize = style?.fontSize ? `${style.fontSize}pt` : '13px';
    const fontFamily =
      style?.fontFamily ?? '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    parts.push(fontSize);
    parts.push(fontFamily);
    return this.textMeasurer.measureWidth(text, parts.join(' ')) + CELL_PADDING * 2;
  }

  private drawBorders(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    w: number,
    h: number,
    style: CellStyle,
  ): void {
    if (style.borderTop) this.drawBorderLine(ctx, x, y, x + w, y, style.borderTop);
    if (style.borderRight) this.drawBorderLine(ctx, x + w, y, x + w, y + h, style.borderRight);
    if (style.borderBottom) this.drawBorderLine(ctx, x, y + h, x + w, y + h, style.borderBottom);
    if (style.borderLeft) this.drawBorderLine(ctx, x, y, x, y + h, style.borderLeft);
  }

  private drawBorderLine(
    ctx: CanvasRenderingContext2D,
    x1: number,
    y1: number,
    x2: number,
    y2: number,
    border: BorderStyle,
  ): void {
    ctx.strokeStyle = border.color;
    ctx.lineWidth = border.width === 'thick' ? 3 : border.width === 'medium' ? 2 : 1;
    if (border.style === 'dashed') ctx.setLineDash([4, 2]);
    else if (border.style === 'dotted') ctx.setLineDash([1, 1]);
    else ctx.setLineDash([]);

    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    ctx.setLineDash([]);
  }

  private drawCellText(
    ctx: CanvasRenderingContext2D,
    x: number,
    y: number,
    w: number,
    h: number,
    text: string,
    rawValue: unknown,
    style?: CellStyle,
  ): void {
    const parts: string[] = [];
    if (style?.italic) parts.push('italic');
    if (style?.bold) parts.push('bold');
    const fontSize = style?.fontSize ? `${style.fontSize}pt` : '13px';
    const fontFamily =
      style?.fontFamily ?? '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif';
    parts.push(fontSize);
    parts.push(fontFamily);
    const font = parts.join(' ');

    ctx.font = font;
    ctx.fillStyle = style?.textColor ?? this.defaultTextColor;

    let textAlign: CanvasTextAlign = 'left';
    if (style?.horizontalAlign) {
      textAlign = style.horizontalAlign as CanvasTextAlign;
    } else if (typeof rawValue === 'number') {
      textAlign = 'right';
    }

    ctx.textAlign = textAlign;
    ctx.textBaseline = 'middle';

    let textY = y + h / 2;
    if (style?.verticalAlign === 'top') textY = y + CELL_PADDING + 6;
    else if (style?.verticalAlign === 'bottom') textY = y + h - CELL_PADDING - 2;

    let textX: number;
    const maxWidth = w - CELL_PADDING * 2;
    if (textAlign === 'center') textX = x + w / 2;
    else if (textAlign === 'right') textX = x + w - CELL_PADDING;
    else textX = x + CELL_PADDING;

    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();

    // Text wrap
    if (style?.wrapText && maxWidth > 0) {
      this.drawWrappedText(ctx, text, textX, y, maxWidth, h, textAlign, font, style);
    } else {
      if (style?.underline) {
        const textWidth = this.textMeasurer.measureWidth(text, font);
        let lineX = textX;
        if (textAlign === 'center') lineX -= textWidth / 2;
        else if (textAlign === 'right') lineX -= textWidth;

        ctx.beginPath();
        ctx.moveTo(lineX, textY + 8);
        ctx.lineTo(lineX + textWidth, textY + 8);
        ctx.strokeStyle = ctx.fillStyle;
        ctx.lineWidth = 1;
        ctx.stroke();
      }
      ctx.fillText(text, textX, textY, maxWidth);
    }

    ctx.restore();
  }

  private drawWrappedText(
    ctx: CanvasRenderingContext2D,
    text: string,
    textX: number,
    cellY: number,
    maxWidth: number,
    cellH: number,
    textAlign: CanvasTextAlign,
    font: string,
    style?: CellStyle,
  ): void {
    const words = text.split(/(\s+)/);
    const lineHeight = 16;
    const lines: string[] = [];
    let currentLine = '';

    for (const word of words) {
      const testLine = currentLine + word;
      const testWidth = this.textMeasurer.measureWidth(testLine, font);
      if (testWidth > maxWidth && currentLine.length > 0) {
        lines.push(currentLine);
        currentLine = word.trimStart();
      } else {
        currentLine = testLine;
      }
    }
    if (currentLine) lines.push(currentLine);

    const totalHeight = lines.length * lineHeight;
    let startY = cellY + (cellH - totalHeight) / 2 + lineHeight / 2;
    if (style?.verticalAlign === 'top') startY = cellY + CELL_PADDING + lineHeight / 2;
    else if (style?.verticalAlign === 'bottom')
      startY = cellY + cellH - totalHeight + lineHeight / 2 - CELL_PADDING;

    for (let i = 0; i < lines.length; i++) {
      ctx.fillText(lines[i], textX, startY + i * lineHeight, maxWidth);
    }
  }

  setDefaultTextColor(color: string): void {
    this.defaultTextColor = color;
  }
}
