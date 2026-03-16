import { describe, it, expect } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import { ContentXmlSerializer } from '../../../src/ods/ContentXmlSerializer';
import { SpreadsheetModel } from '../../../src/model/SpreadsheetModel';

const xmlParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name: string) =>
    ['table:table', 'table:table-row', 'table:table-cell', 'style:style'].includes(name),
};

function parseXml(xml: string): Record<string, unknown> {
  return new XMLParser(xmlParserOptions).parse(xml);
}

function makeModel(): SpreadsheetModel {
  return new SpreadsheetModel();
}

describe('ContentXmlSerializer', () => {
  const serializer = new ContentXmlSerializer();

  describe('serialize() structure', () => {
    it('produces valid XML starting with declaration', () => {
      const xml = serializer.serialize(makeModel());
      expect(xml).toMatch(/^<\?xml/);
    });

    it('contains office:document-content root', () => {
      const xml = serializer.serialize(makeModel());
      expect(xml).toContain('office:document-content');
    });

    it('contains office:body and office:spreadsheet', () => {
      const xml = serializer.serialize(makeModel());
      expect(xml).toContain('office:body');
      expect(xml).toContain('office:spreadsheet');
    });

    it('contains sheet name', () => {
      const model = makeModel();
      model.sheets[0].name = 'MySheet';
      const xml = serializer.serialize(model);
      expect(xml).toContain('MySheet');
    });

    it('serializes multiple sheets', () => {
      const model = makeModel();
      model.addSheet('Sheet2');
      model.addSheet('Sheet3');
      const xml = serializer.serialize(model);
      expect(xml).toContain('Sheet1');
      expect(xml).toContain('Sheet2');
      expect(xml).toContain('Sheet3');
    });
  });

  describe('serialize() cell values', () => {
    it('includes numeric value in output', () => {
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 42, computedValue: 42 });
      const xml = serializer.serialize(model);
      expect(xml).toContain('42');
      expect(xml).toContain('office:value-type="float"');
    });

    it('includes string value in output', () => {
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: 'hello', computedValue: 'hello' });
      const xml = serializer.serialize(model);
      expect(xml).toContain('hello');
      expect(xml).toContain('office:value-type="string"');
    });

    it('includes boolean value in output', () => {
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: true, computedValue: true });
      const xml = serializer.serialize(model);
      expect(xml).toContain('office:value-type="boolean"');
    });

    it('includes formula in output', () => {
      const model = makeModel();
      model.sheets[0].setCell(0, 0, { rawValue: null, formula: 'SUM(A2:A5)', computedValue: null });
      const xml = serializer.serialize(model);
      expect(xml).toContain('table:formula');
      expect(xml).toContain('SUM');
    });

    it('escapes special XML characters in string values', () => {
      const model = makeModel();
      model.sheets[0].setCell(0, 0, {
        rawValue: 'a & b < c > d "e"',
        computedValue: 'a & b < c > d "e"',
      });
      const xml = serializer.serialize(model);
      expect(xml).toContain('&amp;');
      expect(xml).toContain('&lt;');
      expect(xml).not.toContain('a & b');
    });
  });

  describe('serialize() styles', () => {
    it('includes automatic styles section', () => {
      const model = makeModel();
      model.getOrCreateStyle({ id: 'bold-style', bold: true });
      const xml = serializer.serialize(model);
      expect(xml).toContain('office:automatic-styles');
      expect(xml).toContain('bold-style');
    });

    it('includes bold text property', () => {
      const model = makeModel();
      model.getOrCreateStyle({ id: 's1', bold: true });
      const xml = serializer.serialize(model);
      expect(xml).toContain('fo:font-weight="bold"');
    });

    it('includes italic text property', () => {
      const model = makeModel();
      model.getOrCreateStyle({ id: 's1', italic: true });
      const xml = serializer.serialize(model);
      expect(xml).toContain('fo:font-style="italic"');
    });

    it('includes background color', () => {
      const model = makeModel();
      model.getOrCreateStyle({ id: 's1', backgroundColor: '#ff0000' });
      const xml = serializer.serialize(model);
      expect(xml).toContain('#ff0000');
    });

    it('includes horizontal align', () => {
      const model = makeModel();
      model.getOrCreateStyle({ id: 's1', horizontalAlign: 'center' });
      const xml = serializer.serialize(model);
      expect(xml).toContain('center');
    });

    it('includes border properties', () => {
      const model = makeModel();
      model.getOrCreateStyle({
        id: 's1',
        borderTop: { width: 'thin', style: 'solid', color: '#000000' },
      });
      const xml = serializer.serialize(model);
      expect(xml).toContain('fo:border-top');
      expect(xml).toContain('#000000');
    });
  });

  describe('serialize() merge cells', () => {
    it('includes col span attribute for merged cells', () => {
      const model = makeModel();
      model.sheets[0].setCell(0, 0, {
        rawValue: 'merged',
        computedValue: 'merged',
        mergeColSpan: 3,
        mergeRowSpan: 1,
      });
      const xml = serializer.serialize(model);
      expect(xml).toContain('table:number-columns-spanned="3"');
    });

    it('includes row span attribute for merged cells', () => {
      const model = makeModel();
      model.sheets[0].setCell(0, 0, {
        rawValue: 'merged',
        computedValue: 'merged',
        mergeColSpan: 1,
        mergeRowSpan: 2,
      });
      const xml = serializer.serialize(model);
      expect(xml).toContain('table:number-rows-spanned="2"');
    });
  });

  describe('serialize() empty model', () => {
    it('produces parseable XML for empty model', () => {
      const model = makeModel();
      const xml = serializer.serialize(model);
      // Should not throw
      expect(() => parseXml(xml)).not.toThrow();
    });
  });
});
