export const NS = {
  TABLE: 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  TEXT: 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  OFFICE: 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  STYLE: 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  FO: 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  SVG: 'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  OF: 'urn:oasis:names:tc:opendocument:xmlns:of:1.2',
  NUMBER: 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  CALCEXT: 'urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0',
  META: 'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  DC: 'http://purl.org/dc/elements/1.1/',
  XLINK: 'http://www.w3.org/1999/xlink',
  MANIFEST: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
} as const;

/** Prefix used in parsed XML attributes by fast-xml-parser */
export const ATTR = '@_';

/** Common attribute accessor helpers */
export function attr(node: Record<string, unknown>, name: string): string | undefined {
  return node[ATTR + name] as string | undefined;
}

/** Get child array, always returns an array even if single element or missing */
export function children(node: Record<string, unknown>, tag: string): Record<string, unknown>[] {
  const val = node[tag];
  if (!val) return [];
  if (Array.isArray(val)) return val as Record<string, unknown>[];
  return [val as Record<string, unknown>];
}
