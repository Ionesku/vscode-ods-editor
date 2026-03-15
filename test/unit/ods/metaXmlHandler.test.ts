import { describe, it, expect } from 'vitest';
import { XMLParser } from 'fast-xml-parser';
import { MetaXmlHandler } from '../../../src/ods/MetaXmlHandler';

const xmlParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseAttributeValue: false,
  trimValues: false,
};

function parseXml(xml: string): Record<string, unknown> {
  return new XMLParser(xmlParserOptions).parse(xml);
}

describe('MetaXmlHandler', () => {
  const handler = new MetaXmlHandler();

  describe('parse()', () => {
    it('returns empty metadata for empty document', () => {
      const meta = handler.parse({});
      expect(meta).toEqual({});
    });

    it('parses title and description', () => {
      const xml = `<?xml version="1.0"?>
<office:document-meta
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <office:meta>
    <dc:title>My Spreadsheet</dc:title>
    <dc:description>Test file</dc:description>
  </office:meta>
</office:document-meta>`;
      const doc = parseXml(xml);
      const meta = handler.parse(doc);
      expect(meta.title).toBe('My Spreadsheet');
      expect(meta.description).toBe('Test file');
    });

    it('parses creator and creation date', () => {
      const xml = `<?xml version="1.0"?>
<office:document-meta
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/">
  <office:meta>
    <meta:initial-creator>Alice</meta:initial-creator>
    <meta:creation-date>2024-01-15T10:30:00</meta:creation-date>
    <dc:date>2024-03-01T12:00:00</dc:date>
  </office:meta>
</office:document-meta>`;
      const doc = parseXml(xml);
      const meta = handler.parse(doc);
      expect(meta.creator).toBe('Alice');
      expect(meta.creationDate).toBe('2024-01-15T10:30:00');
      expect(meta.modifiedDate).toBe('2024-03-01T12:00:00');
    });

    it('handles missing office:meta gracefully', () => {
      const xml = `<?xml version="1.0"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>`;
      const doc = parseXml(xml);
      const meta = handler.parse(doc);
      expect(meta).toEqual({});
    });
  });

  describe('serialize()', () => {
    it('produces valid XML string', () => {
      const xml = handler.serialize({});
      expect(xml).toContain('<?xml');
      expect(xml).toContain('office:document-meta');
      expect(xml).toContain('office:meta');
    });

    it('includes creator when provided', () => {
      const xml = handler.serialize({ creator: 'Bob' });
      expect(xml).toContain('Bob');
      expect(xml).toContain('meta:initial-creator');
    });

    it('includes title when provided', () => {
      const xml = handler.serialize({ title: 'Budget 2024' });
      expect(xml).toContain('Budget 2024');
      expect(xml).toContain('dc:title');
    });

    it('includes description when provided', () => {
      const xml = handler.serialize({ description: 'Annual report' });
      expect(xml).toContain('Annual report');
      expect(xml).toContain('dc:description');
    });

    it('uses provided creationDate, not current time', () => {
      const xml = handler.serialize({ creationDate: '2020-01-01T00:00:00' });
      expect(xml).toContain('2020-01-01T00:00:00');
    });

    it('escapes special XML characters in creator', () => {
      const xml = handler.serialize({ creator: 'Alice & Bob <test>' });
      expect(xml).toContain('&amp;');
      expect(xml).toContain('&lt;');
      expect(xml).not.toContain('Alice & Bob');
    });

    it('serialize → parse round-trips title and creator', () => {
      const original = { title: 'My Sheet', creator: 'Charlie' };
      const xml = handler.serialize(original);
      const doc = parseXml(xml);
      const parsed = handler.parse(doc);
      expect(parsed.title).toBe(original.title);
      expect(parsed.creator).toBe(original.creator);
    });
  });
});
