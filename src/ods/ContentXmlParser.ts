import { SpreadsheetModel } from '../model/SpreadsheetModel';
import { SheetModel } from '../model/SheetModel';
import { CellData, CellStyle, BorderStyle, createEmptyCell } from '../model/types';
import { attr, children } from './odsNamespaces';

type XNode = Record<string, unknown>;

export class ContentXmlParser {
  parse(doc: XNode, model: SpreadsheetModel): void {
    // Navigate to the spreadsheet body
    const docContent =
      (doc['office:document-content'] as XNode) ?? (doc['office:document'] as XNode);
    if (!docContent) return;

    // Parse automatic styles
    this.parseAutomaticStyles(docContent, model);

    const body = docContent['office:body'] as XNode | undefined;
    if (!body) return;
    const spreadsheet = body['office:spreadsheet'] as XNode | undefined;
    if (!spreadsheet) return;

    // Clear default sheets
    model.sheets = [];

    // Parse sheets
    const tables = children(spreadsheet, 'table:table');
    for (const table of tables) {
      const sheetName = attr(table, 'table:name') ?? 'Sheet';
      const sheet = new SheetModel(sheetName);
      this.parseTable(table, sheet, model);
      model.sheets.push(sheet);
    }

    if (model.sheets.length === 0) {
      model.sheets.push(new SheetModel('Sheet1'));
    }
  }

  private parseAutomaticStyles(docContent: XNode, model: SpreadsheetModel): void {
    const autoStyles = docContent['office:automatic-styles'] as XNode | undefined;
    if (!autoStyles) return;

    const styles = children(autoStyles, 'style:style');
    for (const styleNode of styles) {
      const styleName = attr(styleNode, 'style:name');
      const family = attr(styleNode, 'style:family');
      if (!styleName) continue;

      // Column/row dimension styles — store in rawExtras for parseTable to pick up
      if (family === 'table-column') {
        const colProps = styleNode['style:table-column-properties'] as XNode | undefined;
        if (colProps) {
          const width = attr(colProps, 'style:column-width');
          if (width) {
            const px = this.convertToPixels(width);
            if (px > 0) model.rawExtras.set(`col-width:${styleName}`, String(px));
          }
        }
        continue;
      }
      if (family === 'table-row') {
        const rowProps = styleNode['style:table-row-properties'] as XNode | undefined;
        if (rowProps) {
          const height = attr(rowProps, 'style:row-height');
          if (height) {
            const px = this.convertToPixels(height);
            if (px > 0) model.rawExtras.set(`row-height:${styleName}`, String(px));
          }
        }
        continue;
      }

      if (family !== 'table-cell') continue;

      const cellStyle: CellStyle = { id: styleName };
      const textProps = styleNode['style:text-properties'] as XNode | undefined;
      const cellProps = styleNode['style:table-cell-properties'] as XNode | undefined;
      const paraProps = styleNode['style:paragraph-properties'] as XNode | undefined;

      if (textProps) {
        const fontWeight = attr(textProps, 'fo:font-weight');
        if (fontWeight === 'bold') cellStyle.bold = true;

        const fontStyle = attr(textProps, 'fo:font-style');
        if (fontStyle === 'italic') cellStyle.italic = true;

        const color = attr(textProps, 'fo:color');
        if (color) cellStyle.textColor = color;

        const fontSize = attr(textProps, 'fo:font-size');
        if (fontSize) cellStyle.fontSize = parseFloat(fontSize);

        const fontFamily = attr(textProps, 'style:font-name') ?? attr(textProps, 'fo:font-family');
        if (fontFamily) cellStyle.fontFamily = fontFamily.replace(/'/g, '');

        const underline = attr(textProps, 'style:text-underline-style');
        if (underline && underline !== 'none') cellStyle.underline = true;
      }

      if (cellProps) {
        const bgColor = attr(cellProps, 'fo:background-color');
        if (bgColor && bgColor !== 'transparent') cellStyle.backgroundColor = bgColor;

        const wrapOption = attr(cellProps, 'fo:wrap-option');
        if (wrapOption === 'wrap') cellStyle.wrapText = true;

        const vAlign = attr(cellProps, 'style:vertical-align');
        if (vAlign === 'top' || vAlign === 'middle' || vAlign === 'bottom') {
          cellStyle.verticalAlign = vAlign;
        }

        // Borders
        const border = attr(cellProps, 'fo:border');
        if (border && border !== 'none') {
          const parsed = this.parseBorder(border);
          if (parsed) {
            cellStyle.borderTop = parsed;
            cellStyle.borderRight = parsed;
            cellStyle.borderBottom = parsed;
            cellStyle.borderLeft = parsed;
          }
        }
        // Individual borders override the general border
        const bt = attr(cellProps, 'fo:border-top');
        if (bt && bt !== 'none') cellStyle.borderTop = this.parseBorder(bt) ?? cellStyle.borderTop;
        const br = attr(cellProps, 'fo:border-right');
        if (br && br !== 'none')
          cellStyle.borderRight = this.parseBorder(br) ?? cellStyle.borderRight;
        const bb = attr(cellProps, 'fo:border-bottom');
        if (bb && bb !== 'none')
          cellStyle.borderBottom = this.parseBorder(bb) ?? cellStyle.borderBottom;
        const bl = attr(cellProps, 'fo:border-left');
        if (bl && bl !== 'none')
          cellStyle.borderLeft = this.parseBorder(bl) ?? cellStyle.borderLeft;
      }

      if (paraProps) {
        const align = attr(paraProps, 'fo:text-align');
        if (align === 'start' || align === 'left') cellStyle.horizontalAlign = 'left';
        else if (align === 'center') cellStyle.horizontalAlign = 'center';
        else if (align === 'end' || align === 'right') cellStyle.horizontalAlign = 'right';
      }

      model.styles.set(styleName, cellStyle);
    }
  }

  private parseBorder(value: string): BorderStyle | undefined {
    // Format: "0.06pt solid #000000"
    const parts = value.trim().split(/\s+/);
    if (parts.length < 3) return undefined;

    let width: BorderStyle['width'] = 'thin';
    const size = parseFloat(parts[0]);
    if (size >= 1.5) width = 'thick';
    else if (size >= 0.75) width = 'medium';

    let style: BorderStyle['style'] = 'solid';
    if (parts[1] === 'dashed') style = 'dashed';
    else if (parts[1] === 'dotted') style = 'dotted';

    return { width, style, color: parts[2] };
  }

  private parseTable(table: XNode, sheet: SheetModel, model: SpreadsheetModel): void {
    // Apply column widths from styles parsed by StylesXmlParser
    const columns = children(table, 'table:table-column');
    let colIdx = 0;
    for (const col of columns) {
      const styleName = attr(col, 'table:style-name');
      const repeated = parseInt(attr(col, 'table:number-columns-repeated') ?? '1', 10);
      if (styleName) {
        const pxWidth = Number(model.rawExtras.get(`col-width:${styleName}`));
        if (pxWidth > 0) {
          for (let i = 0; i < repeated; i++) {
            sheet.setColumnWidth(colIdx + i, pxWidth);
          }
        }
      }
      colIdx += repeated;
    }

    // Parse rows, tracking covered positions from multi-row merges
    // Key: "col,row" → origin CellAddress for mergedInto
    const pendingCovered = new Map<string, { col: number; row: number }>();

    let rowIdx = 0;
    const rows = children(table, 'table:table-row');
    for (const row of rows) {
      const rowRepeated = parseInt(attr(row, 'table:number-rows-repeated') ?? '1', 10);

      // Optimization: don't expand massively repeated empty rows
      if (rowRepeated > 1000) {
        const cellNodes = children(row, 'table:table-cell');
        const coveredNodes = children(row, 'table:covered-table-cell');
        const allEmpty = cellNodes.every((c) => this.isCellEmpty(c));
        if (allEmpty && coveredNodes.length === 0) {
          rowIdx += rowRepeated;
          continue;
        }
      }

      // Apply row height from styles
      const rowStyleName = attr(row, 'table:style-name');
      if (rowStyleName) {
        const pxHeight = Number(model.rawExtras.get(`row-height:${rowStyleName}`));
        if (pxHeight > 0) {
          for (let rr = 0; rr < rowRepeated; rr++) {
            sheet.setRowHeight(rowIdx + rr, pxHeight);
          }
        }
      }

      for (let rr = 0; rr < rowRepeated; rr++) {
        this.parseRow(row, sheet, rowIdx, sheet.name, pendingCovered);
        rowIdx++;
      }
    }
  }

  private isCellEmpty(cell: XNode): boolean {
    const valueType = attr(cell, 'office:value-type');
    const formula = attr(cell, 'table:formula');
    const textP = cell['text:p'];
    return !valueType && !formula && !textP;
  }

  private parseRow(
    row: XNode,
    sheet: SheetModel,
    rowIdx: number,
    sheetName: string,
    pendingCovered: Map<string, { col: number; row: number }>,
  ): void {
    let colIdx = 0;
    const cells = children(row, 'table:table-cell');
    // Count of covered cells in this row (from fast-xml-parser separate array)
    const coveredCount = (children(row, 'table:covered-table-cell') as unknown[]).length;
    // Total slots = regular cells (expanded) + covered cells
    // We advance colIdx by skipping pending-covered positions, then by mergeColSpan

    for (const cellNode of cells) {
      const repeated = parseInt(attr(cellNode, 'table:number-columns-repeated') ?? '1', 10);

      // Skip massively repeated empty cells
      if (repeated > 1000 && this.isCellEmpty(cellNode)) {
        colIdx += repeated;
        continue;
      }

      for (let cr = 0; cr < repeated; cr++) {
        // Skip positions that are covered by a preceding multi-row merge
        while (pendingCovered.has(`${colIdx},${rowIdx}`)) {
          const origin = pendingCovered.get(`${colIdx},${rowIdx}`)!;
          sheet.setCell(colIdx, rowIdx, {
            mergedInto: { sheet: sheetName, col: origin.col, row: origin.row },
          });
          colIdx++;
        }

        const cellData = this.parseCell(cellNode, sheetName);
        if (cellData) {
          sheet.setCell(colIdx, rowIdx, cellData);

          // Register covered positions for this cell
          const colSpan = cellData.mergeColSpan ?? 1;
          const rowSpan = cellData.mergeRowSpan ?? 1;
          const originCol = colIdx;
          const originRow = rowIdx;

          // Same-row covered cols: advance colIdx by the full span
          if (colSpan > 1) {
            for (let cc = 1; cc < colSpan; cc++) {
              sheet.setCell(colIdx + cc, rowIdx, {
                mergedInto: { sheet: sheetName, col: originCol, row: originRow },
              });
            }
          }

          // Multi-row covered positions
          if (rowSpan > 1) {
            for (let rr = 1; rr < rowSpan; rr++) {
              for (let cc = 0; cc < colSpan; cc++) {
                pendingCovered.set(`${originCol + cc},${originRow + rr}`, {
                  col: originCol,
                  row: originRow,
                });
              }
            }
          }

          // Advance past same-row covered cells
          colIdx += colSpan;
        } else {
          colIdx++;
        }
      }
    }

    // Consume any remaining covered positions at end of row (needed for correct count)
    void coveredCount; // already handled via pendingCovered logic above
  }

  private parseCell(cellNode: XNode, _sheetName: string): Partial<CellData> | null {
    const valueType = attr(cellNode, 'office:value-type');
    const formula = attr(cellNode, 'table:formula');
    const styleName = attr(cellNode, 'table:style-name');
    const colSpan = parseInt(attr(cellNode, 'table:number-columns-spanned') ?? '1', 10);
    const rowSpan = parseInt(attr(cellNode, 'table:number-rows-spanned') ?? '1', 10);

    // Get text content
    const textP = cellNode['text:p'];
    let textContent: string | null = null;
    if (textP !== undefined) {
      if (typeof textP === 'string' || typeof textP === 'number') {
        textContent = String(textP);
      } else if (typeof textP === 'object' && textP !== null) {
        // Could be an object with #text
        const textObj = textP as Record<string, unknown>;
        if ('#text' in textObj) {
          textContent = String(textObj['#text']);
        } else if (Array.isArray(textP)) {
          // Multiple text:p elements
          textContent = (textP as unknown[])
            .map((t) => {
              if (typeof t === 'string' || typeof t === 'number') return String(t);
              if (
                typeof t === 'object' &&
                t !== null &&
                '#text' in (t as Record<string, unknown>)
              ) {
                return String((t as Record<string, unknown>)['#text']);
              }
              return '';
            })
            .join('\n');
        }
      }
    }

    const hasContent =
      valueType || formula || textContent !== null || styleName || colSpan > 1 || rowSpan > 1;
    if (!hasContent) return null;

    const cell = createEmptyCell();

    // Parse value based on type
    if (valueType === 'float' || valueType === 'currency' || valueType === 'percentage') {
      const val = attr(cellNode, 'office:value');
      cell.rawValue = val !== undefined ? parseFloat(val) : null;
      cell.computedValue = cell.rawValue;
    } else if (valueType === 'string') {
      cell.rawValue = textContent;
      cell.computedValue = textContent;
    } else if (valueType === 'boolean') {
      const val = attr(cellNode, 'office:boolean-value');
      cell.rawValue = val === 'true';
      cell.computedValue = cell.rawValue;
    } else if (valueType === 'date') {
      const val = attr(cellNode, 'office:date-value');
      cell.rawValue = val ?? textContent;
      cell.computedValue = cell.rawValue;
    } else if (textContent !== null) {
      // No value-type but has text
      cell.rawValue = textContent;
      cell.computedValue = textContent;
    }

    // Parse formula
    if (formula) {
      // ODS formulas start with "of:=" or "oooc:="
      let f = formula;
      if (f.startsWith('of:=')) f = f.substring(4);
      else if (f.startsWith('oooc:=')) f = f.substring(6);
      else if (f.startsWith('=')) f = f.substring(1);
      cell.formula = f;
    }

    // Style
    if (styleName) {
      cell.styleId = styleName;
    }

    // Merge
    if (colSpan > 1 || rowSpan > 1) {
      cell.mergeColSpan = colSpan;
      cell.mergeRowSpan = rowSpan;
    }

    return cell;
  }

  /** Convert ODF length units to pixels (same logic as StylesXmlParser) */
  private convertToPixels(value: string): number {
    const num = parseFloat(value);
    if (isNaN(num)) return 0;
    if (value.endsWith('in')) return num * 96;
    if (value.endsWith('cm')) return num * 37.795275591;
    if (value.endsWith('mm')) return num * 3.7795275591;
    if (value.endsWith('pt')) return num * 1.333;
    if (value.endsWith('px')) return num;
    return num * 96;
  }
}
