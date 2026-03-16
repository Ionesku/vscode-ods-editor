import { SpreadsheetModel } from '../model/SpreadsheetModel';
import { SheetModel, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '../model/SheetModel';
import { CellData, CellStyle, CellValue } from '../model/types';

/** Convert pixels to centimetres (ODS unit), rounded to 6 decimal places */
function pxToCm(px: number): string {
  return (px / 37.795275591).toFixed(6) + 'cm';
}

export class ContentXmlSerializer {
  serialize(model: SpreadsheetModel): string {
    let xml = `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
  xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
  office:version="1.2">`;

    // Collect unique non-default column widths and row heights across all sheets
    // so we can emit shared automatic styles for them.
    const colWidthStyles = new Map<number, string>(); // px → style name
    const rowHeightStyles = new Map<number, string>(); // px → style name

    for (const sheet of model.sheets) {
      const usedRange = sheet.getUsedRange();
      const maxCol = usedRange ? usedRange.endCol + 1 : 1;
      const maxRow = usedRange ? usedRange.endRow + 1 : 0;

      for (let c = 0; c < maxCol; c++) {
        const w = sheet.getColumnWidth(c);
        if (w !== DEFAULT_COL_WIDTH && !colWidthStyles.has(w)) {
          colWidthStyles.set(w, `co${colWidthStyles.size + 1}`);
        }
      }
      for (let r = 0; r < maxRow; r++) {
        const h = sheet.getRowHeight(r);
        if (h !== DEFAULT_ROW_HEIGHT && !rowHeightStyles.has(h)) {
          rowHeightStyles.set(h, `ro${rowHeightStyles.size + 1}`);
        }
      }
    }

    // Automatic styles
    xml += '\n  <office:automatic-styles>';
    for (const [, style] of model.styles) {
      xml += this.serializeStyle(style);
    }
    // Column width styles
    for (const [px, name] of colWidthStyles) {
      xml += `\n    <style:style style:name="${name}" style:family="table-column">`;
      xml += `\n      <style:table-column-properties style:column-width="${pxToCm(px)}"/>`;
      xml += '\n    </style:style>';
    }
    // Default column style (referenced for default-width columns)
    xml += `\n    <style:style style:name="co-default" style:family="table-column">`;
    xml += `\n      <style:table-column-properties style:column-width="${pxToCm(DEFAULT_COL_WIDTH)}"/>`;
    xml += '\n    </style:style>';
    // Row height styles
    for (const [px, name] of rowHeightStyles) {
      xml += `\n    <style:style style:name="${name}" style:family="table-row">`;
      xml += `\n      <style:table-row-properties style:row-height="${pxToCm(px)}" style:use-optimal-row-height="false"/>`;
      xml += '\n    </style:style>';
    }
    xml += '\n  </office:automatic-styles>';

    // Body
    xml += '\n  <office:body>\n    <office:spreadsheet>';
    for (const sheet of model.sheets) {
      xml += this.serializeSheet(sheet, colWidthStyles, rowHeightStyles);
    }
    xml += '\n    </office:spreadsheet>\n  </office:body>';
    xml += '\n</office:document-content>';
    return xml;
  }

  private serializeStyle(style: CellStyle): string {
    let xml = `\n    <style:style style:name="${esc(style.id)}" style:family="table-cell">`;

    // Text properties
    const textParts: string[] = [];
    if (style.bold) textParts.push('fo:font-weight="bold"');
    if (style.italic) textParts.push('fo:font-style="italic"');
    if (style.textColor) textParts.push(`fo:color="${esc(style.textColor)}"`);
    if (style.fontSize) textParts.push(`fo:font-size="${style.fontSize}pt"`);
    if (style.fontFamily) textParts.push(`style:font-name="${esc(style.fontFamily)}"`);
    if (style.underline) textParts.push('style:text-underline-style="solid"');
    if (style.strikethrough) textParts.push('style:text-line-through-style="solid"');
    if (textParts.length > 0) {
      xml += `\n      <style:text-properties ${textParts.join(' ')}/>`;
    }

    // Cell properties
    const cellParts: string[] = [];
    if (style.backgroundColor)
      cellParts.push(`fo:background-color="${esc(style.backgroundColor)}"`);
    if (style.wrapText) cellParts.push('fo:wrap-option="wrap"');
    if (style.verticalAlign) cellParts.push(`style:vertical-align="${style.verticalAlign}"`);
    // Borders
    if (style.borderTop) cellParts.push(`fo:border-top="${this.borderToString(style.borderTop)}"`);
    if (style.borderRight)
      cellParts.push(`fo:border-right="${this.borderToString(style.borderRight)}"`);
    if (style.borderBottom)
      cellParts.push(`fo:border-bottom="${this.borderToString(style.borderBottom)}"`);
    if (style.borderLeft)
      cellParts.push(`fo:border-left="${this.borderToString(style.borderLeft)}"`);
    if (cellParts.length > 0) {
      xml += `\n      <style:table-cell-properties ${cellParts.join(' ')}/>`;
    }

    // Paragraph properties
    if (style.horizontalAlign) {
      const align =
        style.horizontalAlign === 'left'
          ? 'start'
          : style.horizontalAlign === 'right'
            ? 'end'
            : 'center';
      xml += `\n      <style:paragraph-properties fo:text-align="${align}"/>`;
    }

    xml += '\n    </style:style>';
    return xml;
  }

  private borderToString(border: import('../model/types').BorderStyle): string {
    const widthMap = { thin: '0.06pt', medium: '1pt', thick: '2pt' };
    return `${widthMap[border.width]} ${border.style} ${border.color}`;
  }

  private serializeSheet(
    sheet: SheetModel,
    colWidthStyles: Map<number, string>,
    rowHeightStyles: Map<number, string>,
  ): string {
    let xml = `\n      <table:table table:name="${esc(sheet.name)}">`;

    // Columns — group consecutive columns with the same width to reduce output size
    const usedRange = sheet.getUsedRange();
    const maxCol = usedRange ? usedRange.endCol + 1 : 1;
    let c = 0;
    while (c < maxCol) {
      const w = sheet.getColumnWidth(c);
      const styleName = colWidthStyles.get(w) ?? 'co-default';
      let repeat = 1;
      while (c + repeat < maxCol && sheet.getColumnWidth(c + repeat) === w) {
        repeat++;
      }
      const repeatAttr = repeat > 1 ? ` table:number-columns-repeated="${repeat}"` : '';
      xml += `\n        <table:table-column table:style-name="${styleName}"${repeatAttr}/>`;
      c += repeat;
    }

    // Rows — track covered positions to emit <table:covered-table-cell/> correctly
    const pendingCoveredRows = new Map<number, Set<number>>(); // row → Set<col>
    const maxRow = usedRange ? usedRange.endRow + 1 : 0;
    for (let r = 0; r < maxRow; r++) {
      const coveredInRow: Set<number> = pendingCoveredRows.get(r) ?? new Set();
      const h = sheet.getRowHeight(r);
      const rowStyleAttr = rowHeightStyles.has(h)
        ? ` table:style-name="${rowHeightStyles.get(h)}"`
        : '';
      xml += `\n        <table:table-row${rowStyleAttr}>`;
      for (let col = 0; col < maxCol; col++) {
        if (coveredInRow.has(col) || sheet.getCell(col, r).mergedInto !== null) {
          xml += '\n          <table:covered-table-cell/>';
        } else {
          const cell = sheet.getCell(col, r);
          xml += this.serializeCell(cell, sheet.name);
          // Mark same-row covered positions from this origin cell
          if (cell.mergeColSpan > 1) {
            for (let cc = 1; cc < cell.mergeColSpan; cc++) {
              coveredInRow.add(col + cc);
            }
          }
          // Register multi-row covered positions
          if (cell.mergeRowSpan > 1) {
            for (let rr = 1; rr < cell.mergeRowSpan; rr++) {
              for (let cc = 0; cc < cell.mergeColSpan; cc++) {
                if (!pendingCoveredRows.has(r + rr)) pendingCoveredRows.set(r + rr, new Set());
                pendingCoveredRows.get(r + rr)!.add(col + cc);
              }
            }
          }
        }
      }
      xml += '\n        </table:table-row>';
    }

    xml += '\n      </table:table>';
    return xml;
  }

  private serializeCell(cell: CellData, _sheetName: string): string {
    const attrs: string[] = [];

    // Style
    if (cell.styleId) {
      attrs.push(`table:style-name="${esc(cell.styleId)}"`);
    }

    // Formula
    if (cell.formula) {
      attrs.push(`table:formula="of:=${esc(cell.formula)}"`);
    }

    // Merge
    if (cell.mergeColSpan > 1) {
      attrs.push(`table:number-columns-spanned="${cell.mergeColSpan}"`);
    }
    if (cell.mergeRowSpan > 1) {
      attrs.push(`table:number-rows-spanned="${cell.mergeRowSpan}"`);
    }

    // Value type and value
    const value = cell.computedValue ?? cell.rawValue;
    if (
      value === null &&
      !cell.formula &&
      !cell.styleId &&
      cell.mergeColSpan === 1 &&
      cell.mergeRowSpan === 1
    ) {
      return '\n          <table:table-cell/>';
    }

    if (typeof value === 'number') {
      attrs.push('office:value-type="float"');
      attrs.push(`office:value="${value}"`);
    } else if (typeof value === 'boolean') {
      attrs.push('office:value-type="boolean"');
      attrs.push(`office:boolean-value="${value}"`);
    } else if (typeof value === 'string') {
      attrs.push('office:value-type="string"');
    }

    const attrStr = attrs.length > 0 ? ' ' + attrs.join(' ') : '';
    const displayValue = this.formatDisplayValue(value);

    const commentXml = cell.comment
      ? `<office:annotation><text:p>${esc(cell.comment)}</text:p></office:annotation>`
      : '';

    if (displayValue !== null) {
      return `\n          <table:table-cell${attrStr}>${commentXml}<text:p>${esc(displayValue)}</text:p></table:table-cell>`;
    }
    if (commentXml) {
      return `\n          <table:table-cell${attrStr}>${commentXml}</table:table-cell>`;
    }
    return `\n          <table:table-cell${attrStr}/>`;
  }

  private formatDisplayValue(value: CellValue): string | null {
    if (value === null) return null;
    if (typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value === 'string') return value;
    // FormulaError
    return String(value);
  }
}

function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
