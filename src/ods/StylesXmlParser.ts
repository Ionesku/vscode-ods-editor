import { SpreadsheetModel } from '../model/SpreadsheetModel';
import { attr, children } from './odsNamespaces';

type XNode = Record<string, unknown>;

export class StylesXmlParser {
  parse(doc: XNode, model: SpreadsheetModel): void {
    const docStyles = (doc['office:document-styles'] as XNode) ?? (doc['office:document'] as XNode);
    if (!docStyles) return;

    // Parse default column/row styles for dimension info
    const autoStyles = docStyles['office:automatic-styles'] as XNode | undefined;
    if (autoStyles) {
      this.parseColumnRowStyles(autoStyles, model);
    }

    const styles = docStyles['office:styles'] as XNode | undefined;
    if (styles) {
      this.parseColumnRowStyles(styles, model);
    }
  }

  private parseColumnRowStyles(container: XNode, model: SpreadsheetModel): void {
    const styles = children(container, 'style:style');
    for (const styleNode of styles) {
      const family = attr(styleNode, 'style:family');
      const styleName = attr(styleNode, 'style:name');
      if (!styleName) continue;

      if (family === 'table-column') {
        const colProps = styleNode['style:table-column-properties'] as XNode | undefined;
        if (colProps) {
          const width = attr(colProps, 'style:column-width');
          if (width) {
            // Store for later application to sheets
            const pxWidth = this.convertToPixels(width);
            if (pxWidth > 0) {
              model.rawExtras.set(`col-width:${styleName}`, String(pxWidth));
            }
          }
        }
      } else if (family === 'table-row') {
        const rowProps = styleNode['style:table-row-properties'] as XNode | undefined;
        if (rowProps) {
          const height = attr(rowProps, 'style:row-height');
          if (height) {
            const pxHeight = this.convertToPixels(height);
            if (pxHeight > 0) {
              model.rawExtras.set(`row-height:${styleName}`, String(pxHeight));
            }
          }
        }
      }
    }
  }

  /** Convert ODF length units to pixels (approximate) */
  private convertToPixels(value: string): number {
    const num = parseFloat(value);
    if (isNaN(num)) return 0;

    if (value.endsWith('in')) return num * 96;
    if (value.endsWith('cm')) return num * 37.8;
    if (value.endsWith('mm')) return num * 3.78;
    if (value.endsWith('pt')) return num * 1.333;
    if (value.endsWith('px')) return num;
    // Default: assume inches for backwards compat
    return num * 96;
  }
}
