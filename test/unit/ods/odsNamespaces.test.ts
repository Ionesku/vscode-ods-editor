import { describe, it, expect } from 'vitest';
import { attr, children, NS, ATTR } from '../../../src/ods/odsNamespaces';

describe('odsNamespaces', () => {
  describe('NS', () => {
    it('has expected namespace URIs', () => {
      expect(NS.TABLE).toContain('table');
      expect(NS.OFFICE).toContain('office');
      expect(NS.STYLE).toContain('style');
      expect(NS.FO).toContain('xsl-fo');
    });
  });

  describe('ATTR', () => {
    it('is "@_" (fast-xml-parser prefix)', () => {
      expect(ATTR).toBe('@_');
    });
  });

  describe('attr()', () => {
    it('returns the attribute value', () => {
      const node = { '@_table:name': 'Sheet1' };
      expect(attr(node, 'table:name')).toBe('Sheet1');
    });

    it('returns undefined for missing attribute', () => {
      expect(attr({}, 'table:name')).toBeUndefined();
    });

    it('returns undefined when node is empty', () => {
      expect(attr({}, 'anything')).toBeUndefined();
    });

    it('handles various attribute names', () => {
      const node = {
        '@_office:value-type': 'float',
        '@_office:value': '42',
        '@_style:name': 'ce1',
      };
      expect(attr(node, 'office:value-type')).toBe('float');
      expect(attr(node, 'office:value')).toBe('42');
      expect(attr(node, 'style:name')).toBe('ce1');
    });
  });

  describe('children()', () => {
    it('returns array when value is already an array', () => {
      const node = { 'table:table-row': [{ a: 1 }, { a: 2 }] };
      const result = children(node, 'table:table-row');
      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({ a: 1 });
    });

    it('wraps single object in array', () => {
      const node = { 'table:table-cell': { value: 'x' } };
      const result = children(node, 'table:table-cell');
      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({ value: 'x' });
    });

    it('returns empty array for missing key', () => {
      const result = children({}, 'table:table-row');
      expect(result).toEqual([]);
    });

    it('returns empty array when value is falsy', () => {
      const node = { 'table:table-row': null };
      const result = children(node as Record<string, unknown>, 'table:table-row');
      expect(result).toEqual([]);
    });
  });
});
