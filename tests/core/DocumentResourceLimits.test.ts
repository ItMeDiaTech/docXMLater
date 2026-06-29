/**
 * Issue #30 — resource-limit enforcement reachable from the stable Document API.
 *
 * `ResourceLimitError` and `SizeLimitOptions` are deliberately exported from the
 * stable package entry ('../../src'), unlike the rest of the DocxError hierarchy
 * that lives on 'docxmlater/internal'. That lets untrusted-input callers catch
 * the typed error by class and tune the limits without reaching into internals.
 *
 * These tests prove:
 *  1. `ResourceLimitError` is importable from the stable entry and is thrown by
 *     `Document.loadFromBuffer` when a configured budget is breached.
 *  2. A normal small document still loads with the generous defaults.
 *  3. `Document.loadFromBase64` now accepts `sizeLimits` (its options parameter
 *     was widened to `DocumentLoadOptions`) and rejects the same way — this test
 *     would not compile if the parameter were still `DocumentOptions`.
 *
 * The archives are built in-memory with JSZip and are never weaponized payloads —
 * just enough entries to trip a deliberately tiny `maxEntryCount` budget.
 */

import JSZip from 'jszip';
import { Document, Paragraph, ResourceLimitError } from '../../src';

/**
 * Builds a structurally-valid DOCX (the three required parts) plus `extra`
 * additional dummy entries, so the archive's file count exceeds a small
 * configured `maxEntryCount`.
 */
async function buildMultiEntryDocx(extra: number): Promise<Buffer> {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>'
  );
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ' +
      'Target="word/document.xml"/></Relationships>'
  );
  zip.file(
    'word/document.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'
  );
  for (let i = 0; i < extra; i++) {
    zip.file(`word/extra-${i}.xml`, `<x>${i}</x>`);
  }
  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('Document resource limits via stable API (issue #30)', () => {
  it('loadFromBuffer rejects with ResourceLimitError when maxEntryCount is exceeded', async () => {
    // 3 required parts + 5 extra = 8 file entries; budget of 4 must reject.
    const buffer = await buildMultiEntryDocx(5);

    await expect(
      Document.loadFromBuffer(buffer, { sizeLimits: { maxEntryCount: 4 } })
    ).rejects.toThrow(ResourceLimitError);
  });

  it('loads a normal small document with the generous defaults', async () => {
    const source = Document.create();
    const para = new Paragraph();
    para.addText('Defaults should not reject this.');
    source.addParagraph(para);
    const buffer = await source.toBuffer();
    source.dispose();

    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.toPlainText()).toContain('Defaults should not reject this.');
    doc.dispose();
  });

  it('loadFromBase64 accepts sizeLimits and rejects with ResourceLimitError', async () => {
    // Proves loadFromBase64's options widened to DocumentLoadOptions: passing
    // `sizeLimits` here would be a compile error against the old DocumentOptions.
    const buffer = await buildMultiEntryDocx(5);
    const base64 = buffer.toString('base64');

    await expect(
      Document.loadFromBase64(base64, { sizeLimits: { maxEntryCount: 4 } })
    ).rejects.toThrow(ResourceLimitError);
  });
});
