export const ROW_HEADER_WIDTH = 50;
export const COL_HEADER_HEIGHT = 24;

export interface VisibleRange {
  startCol: number;
  endCol: number;
  startRow: number;
  endRow: number;
}

export class ViewportCalculator {
  private colPrefixSums: Float64Array = new Float64Array(0);
  private rowPrefixSums: Float64Array = new Float64Array(0);
  private _totalWidth = 0;
  private _totalHeight = 0;

  updateColumnWidths(widths: number[]): void {
    this.colPrefixSums = new Float64Array(widths.length + 1);
    for (let i = 0; i < widths.length; i++) {
      this.colPrefixSums[i + 1] = this.colPrefixSums[i] + widths[i];
    }
    this._totalWidth = this.colPrefixSums[widths.length];
  }

  updateRowHeights(heights: number[]): void {
    this.rowPrefixSums = new Float64Array(heights.length + 1);
    for (let i = 0; i < heights.length; i++) {
      this.rowPrefixSums[i + 1] = this.rowPrefixSums[i] + heights[i];
    }
    this._totalHeight = this.rowPrefixSums[heights.length];
  }

  get totalWidth(): number {
    return this._totalWidth;
  }
  get totalHeight(): number {
    return this._totalHeight;
  }

  /** Get the X pixel offset of a column's left edge */
  colLeft(col: number): number {
    if (col < 0 || col >= this.colPrefixSums.length) return 0;
    return this.colPrefixSums[col];
  }

  /** Get the width of a column */
  colWidth(col: number): number {
    if (col < 0 || col + 1 >= this.colPrefixSums.length) return 0;
    return this.colPrefixSums[col + 1] - this.colPrefixSums[col];
  }

  /** Get the Y pixel offset of a row's top edge */
  rowTop(row: number): number {
    if (row < 0 || row >= this.rowPrefixSums.length) return 0;
    return this.rowPrefixSums[row];
  }

  /** Get the height of a row */
  rowHeight(row: number): number {
    if (row < 0 || row + 1 >= this.rowPrefixSums.length) return 0;
    return this.rowPrefixSums[row + 1] - this.rowPrefixSums[row];
  }

  /** Binary search: pixel X -> column index */
  colAtX(x: number): number {
    return this.binarySearch(this.colPrefixSums, x);
  }

  /** Binary search: pixel Y -> row index */
  rowAtY(y: number): number {
    return this.binarySearch(this.rowPrefixSums, y);
  }

  /** Get visible cell range for the current viewport */
  getVisibleRange(scrollX: number, scrollY: number, viewW: number, viewH: number): VisibleRange {
    const contentW = viewW - ROW_HEADER_WIDTH;
    const contentH = viewH - COL_HEADER_HEIGHT;

    const startCol = Math.max(0, this.colAtX(scrollX));
    const endCol = Math.min(this.colPrefixSums.length - 2, this.colAtX(scrollX + contentW) + 1);
    const startRow = Math.max(0, this.rowAtY(scrollY));
    const endRow = Math.min(this.rowPrefixSums.length - 2, this.rowAtY(scrollY + contentH) + 1);

    return { startCol, endCol, startRow, endRow };
  }

  private binarySearch(arr: Float64Array, value: number): number {
    let lo = 0;
    let hi = arr.length - 2;
    while (lo <= hi) {
      const mid = (lo + hi) >>> 1;
      if (arr[mid + 1] <= value) {
        lo = mid + 1;
      } else if (arr[mid] > value) {
        hi = mid - 1;
      } else {
        return mid;
      }
    }
    return Math.max(0, lo);
  }
}
