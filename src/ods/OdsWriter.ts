import JSZip from 'jszip';
import { SpreadsheetModel } from '../model/SpreadsheetModel';
import { ContentXmlSerializer } from './ContentXmlSerializer';
import { MetaXmlHandler } from './MetaXmlHandler';

export class OdsWriter {
  async write(model: SpreadsheetModel): Promise<Uint8Array> {
    const zip = new JSZip();

    // mimetype must be first and uncompressed
    zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', {
      compression: 'STORE',
    });

    // content.xml
    const contentXml = new ContentXmlSerializer().serialize(model);
    zip.file('content.xml', contentXml);

    // styles.xml
    zip.file('styles.xml', this.generateStylesXml());

    // meta.xml
    const metaXml = new MetaXmlHandler().serialize(model.metadata);
    zip.file('meta.xml', metaXml);

    // META-INF/manifest.xml
    zip.file('META-INF/manifest.xml', this.generateManifest());

    // settings.xml (minimal)
    zip.file('settings.xml', this.generateSettings());

    return zip.generateAsync({ type: 'uint8array' });
  }

  private generateStylesXml(): string {
    return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.2">
  <office:styles>
    <style:default-style style:family="table-cell">
      <style:text-properties fo:font-size="10pt" style:font-name="Arial"/>
    </style:default-style>
  </office:styles>
</office:document-styles>`;
  }

  private generateManifest(): string {
    return `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
</manifest:manifest>`;
  }

  private generateSettings(): string {
    return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  office:version="1.2">
  <office:settings/>
</office:document-settings>`;
  }
}
