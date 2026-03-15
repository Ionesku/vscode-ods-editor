export class TextMeasurer {
  private cache = new Map<string, number>();
  private ctx: CanvasRenderingContext2D;

  constructor(ctx: CanvasRenderingContext2D) {
    this.ctx = ctx;
  }

  measureWidth(text: string, font: string): number {
    const key = font + '|' + text;
    let width = this.cache.get(key);
    if (width !== undefined) return width;

    this.ctx.font = font;
    width = this.ctx.measureText(text).width;
    this.cache.set(key, width);

    // Prevent cache from growing too large
    if (this.cache.size > 10000) {
      const entries = Array.from(this.cache.entries());
      this.cache = new Map(entries.slice(entries.length - 5000));
    }

    return width;
  }

  clear(): void {
    this.cache.clear();
  }
}
