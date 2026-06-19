/**
 * When parseSDTFromObject returns null (an SDT the framework cannot model —
 * e.g. an unexpected failure in SDT-property parsing), the body-parse loop
 * used to skip the whole w:sdt block without storing anything, so the content
 * control and every paragraph/table inside it were deleted on the next save.
 * It now degrades to a PreservedElement, mirroring the failed
 * registered-element path, so the raw SDT XML round-trips.
 */

import { Document } from '../../src/core/Document';
import { DocumentParser } from '../../src/core/DocumentParser';
import { ZipHandler } from '../../src/zip/ZipHandler';

const BLOCK_SDT =
  `<w:sdt>` +
  `<w:sdtPr><w:id w:val="424242"/><w:tag w:val="ImportantControl"/></w:sdtPr>` +
  `<w:sdtContent>` +
  `<w:p><w:r><w:t>SDT-INNER-CONTENT</w:t></w:r></w:p>` +
  `</w:sdtContent>` +
  `</w:sdt>`;

async function buildDocxWithBlockSdt(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('SDT fallback round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  const docXml = zip.getFileAsString('word/document.xml')!;
  const updated = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${BLOCK_SDT}<w:sectPr`)
    : docXml.replace('</w:body>', `${BLOCK_SDT}</w:body>`);
  zip.updateFile('word/document.xml', updated);

  return zip.toBuffer();
}

describe('Failed SDT parse degrades to PreservedElement (C66)', () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('preserves the raw w:sdt block when parseSDTFromObject returns null', async () => {
    // Force the "cannot model this control" path so the fallback is exercised.
    jest
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .spyOn(DocumentParser.prototype as any, 'parseSDTFromObject')
      .mockResolvedValue(null);

    const buffer = await buildDocxWithBlockSdt();
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const outZip = new ZipHandler();
    await outZip.loadFromBuffer(out);
    const outXml = outZip.getFileAsString('word/document.xml') ?? '';

    // The whole content control plus its inner content must survive the save.
    expect(outXml).toContain('<w:sdt>');
    expect(outXml).toContain('w:val="424242"');
    expect(outXml).toContain('SDT-INNER-CONTENT');
  });
});
