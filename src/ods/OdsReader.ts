import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { SpreadsheetModel } from '../model/SpreadsheetModel';
import { ContentXmlParser } from './ContentXmlParser';
import { StylesXmlParser } from './StylesXmlParser';
import { MetaXmlHandler } from './MetaXmlHandler';

const xmlParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  trimValues: false,
  preserveOrder: false,
  isArray: (name: string) => {
    // These elements can appear multiple times
    return [
      'table:table',
      'table:table-row',
      'table:table-cell',
      'table:covered-table-cell',
      'table:table-column',
      'text:p',
      'style:style',
    ].includes(name);
  },
};

export class OdsReader {
  async read(buffer: Uint8Array): Promise<SpreadsheetModel> {
    const zip = await JSZip.loadAsync(buffer);
    const model = new SpreadsheetModel();
    model.sheets = [];

    const parser = new XMLParser(xmlParserOptions);

    // Parse content.xml (required)
    const contentFile = zip.file('content.xml');
    if (contentFile) {
      const contentXml = await contentFile.async('string');
      const contentDoc = parser.parse(contentXml);
      new ContentXmlParser().parse(contentDoc, model);
    }

    // Parse styles.xml (optional)
    const stylesFile = zip.file('styles.xml');
    if (stylesFile) {
      const stylesXml = await stylesFile.async('string');
      const stylesDoc = parser.parse(stylesXml);
      new StylesXmlParser().parse(stylesDoc, model);
    }

    // Parse meta.xml (optional)
    const metaFile = zip.file('meta.xml');
    if (metaFile) {
      const metaXml = await metaFile.async('string');
      const metaDoc = parser.parse(metaXml);
      model.metadata = new MetaXmlHandler().parse(metaDoc);
    }

    // Ensure at least one sheet
    if (model.sheets.length === 0) {
      model.addSheet('Sheet1');
    }

    return model;
  }
}
