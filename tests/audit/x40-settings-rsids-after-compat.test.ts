/**
 * Regression: generateSettings() must emit <w:rsids> after <w:compat>.
 * CT_Settings is an xsd:sequence where w:compat (#81) precedes w:rsids (#83);
 * the from-scratch settings template previously emitted rsids first, producing
 * schema-invalid settings.xml when RSIDs are set on a new document. The
 * loaded-document merge path (mergeRsidsIntoSettings) already inserts rsids
 * after </w:compat>.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

describe('generateSettings() CT_Settings element order for w:rsids', () => {
  it('places w:rsids after w:compat in a new document with RSIDs set', async () => {
    const doc = Document.create();
    try {
      doc.createParagraph('Body');
      doc.setRsidRoot('00123456');
      doc.addRsid('00ABCDEF');
      const buffer = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buffer);
      const settings = out.getFileAsString('word/settings.xml')!;

      expect(settings).toContain('<w:rsidRoot w:val="00123456"/>');
      expect(settings).toContain('<w:rsid w:val="00ABCDEF"/>');

      const compatClose = settings.indexOf('</w:compat>');
      const rsidsOpen = settings.indexOf('<w:rsids>');
      expect(compatClose).toBeGreaterThan(-1);
      expect(rsidsOpen).toBeGreaterThan(compatClose);

      // rsids must still precede the trailing themeFontLang element
      const themeFontLang = settings.indexOf('<w:themeFontLang');
      expect(themeFontLang).toBeGreaterThan(settings.indexOf('</w:rsids>'));
    } finally {
      doc.dispose();
    }
  });
});
