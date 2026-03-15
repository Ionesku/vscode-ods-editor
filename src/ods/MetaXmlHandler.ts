import { DocumentMetadata } from '../model/types';
import { children } from './odsNamespaces';

type XNode = Record<string, unknown>;

export class MetaXmlHandler {
  parse(doc: XNode): DocumentMetadata {
    const meta: DocumentMetadata = {};
    const docMeta = (doc['office:document-meta'] as XNode) ?? (doc['office:document'] as XNode);
    if (!docMeta) return meta;

    const officeMeta = docMeta['office:meta'] as XNode | undefined;
    if (!officeMeta) return meta;

    if (officeMeta['dc:title']) {
      meta.title = String(officeMeta['dc:title']);
    }
    if (officeMeta['dc:description']) {
      meta.description = String(officeMeta['dc:description']);
    }
    if (officeMeta['meta:initial-creator']) {
      meta.creator = String(officeMeta['meta:initial-creator']);
    }
    if (officeMeta['meta:creation-date']) {
      meta.creationDate = String(officeMeta['meta:creation-date']);
    }
    if (officeMeta['dc:date']) {
      meta.modifiedDate = String(officeMeta['dc:date']);
    }

    return meta;
  }

  serialize(metadata: DocumentMetadata): string {
    const now = new Date().toISOString();
    return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  office:version="1.2">
  <office:meta>
    <meta:creation-date>${metadata.creationDate ?? now}</meta:creation-date>
    <dc:date>${now}</dc:date>${metadata.creator ? `\n    <meta:initial-creator>${escapeXml(metadata.creator)}</meta:initial-creator>` : ''}${metadata.title ? `\n    <dc:title>${escapeXml(metadata.title)}</dc:title>` : ''}${metadata.description ? `\n    <dc:description>${escapeXml(metadata.description)}</dc:description>` : ''}
  </office:meta>
</office:document-meta>`;
  }
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
