/**
 * Inline structured document tags (CT_SdtRun, ECMA-376 §17.5.2.31) are legal
 * children of w:p — date pickers, dropdowns, citations, plain/rich-text
 * content controls inline in a sentence. They have no run-level editing
 * model, so the parser must preserve them verbatim; dropping them deletes
 * the user-visible text inside the control on every load → save cycle.
 */
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

const INLINE_SDT_PARAGRAPH =
  `<w:p>` +
  `<w:r><w:t>BEFORE-</w:t></w:r>` +
  `<w:sdt>` +
  `<w:sdtPr><w:id w:val="123456"/><w:text/></w:sdtPr>` +
  `<w:sdtContent><w:r><w:t>SDT-TEXT</w:t></w:r></w:sdtContent>` +
  `</w:sdt>` +
  `<w:r><w:t>-AFTER</w:t></w:r>` +
  `</w:p>`;

async function buildDocxWithInlineSdt(): Promise<Buffer> {
  const seed = Document.create();
  seed.createParagraph('Inline SDT round-trip');
  const base = await seed.toBuffer();
  seed.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(base);

  const docXml = zip.getFileAsString('word/document.xml')!;
  const updatedDoc = docXml.includes('<w:sectPr')
    ? docXml.replace('<w:sectPr', `${INLINE_SDT_PARAGRAPH}<w:sectPr`)
    : docXml.replace('</w:body>', `${INLINE_SDT_PARAGRAPH}</w:body>`);
  zip.updateFile('word/document.xml', updatedDoc);

  return zip.toBuffer();
}

describe('Inline w:sdt (run-level content control) round-trip', () => {
  it('preserves the w:sdt wrapper and its inner text on unmodified round-trip', async () => {
    const buf1 = await buildDocxWithInlineSdt();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const docXml = out.getFileAsString('word/document.xml')!;
      expect(docXml).toContain('<w:sdt>');
      expect(docXml).toContain('<w:sdtContent>');
      expect(docXml).toContain('SDT-TEXT');
    } finally {
      doc.dispose();
    }
  });

  it('keeps sibling runs intact and in order around the inline SDT', async () => {
    const buf1 = await buildDocxWithInlineSdt();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      const buf2 = await doc.toBuffer();

      const out = new ZipHandler();
      await out.loadFromBuffer(buf2);
      const docXml = out.getFileAsString('word/document.xml')!;
      const beforePos = docXml.indexOf('BEFORE-');
      const sdtPos = docXml.indexOf('<w:sdt>');
      const afterPos = docXml.indexOf('-AFTER');
      expect(beforePos).toBeGreaterThan(-1);
      expect(sdtPos).toBeGreaterThan(beforePos);
      expect(afterPos).toBeGreaterThan(sdtPos);
    } finally {
      doc.dispose();
    }
  });

  it('does not shift run indices when sdtContent contains nested runs', async () => {
    const buf1 = await buildDocxWithInlineSdt();
    const doc = await Document.loadFromBuffer(buf1);
    try {
      // The scanner must skip the whole w:sdt subtree; counting the nested
      // run would misalign indices and duplicate or drop sibling runs.
      const text = doc.getParagraphs()[1]!.getText();
      expect(text).toContain('BEFORE-');
      expect(text).toContain('-AFTER');
    } finally {
      doc.dispose();
    }
  });
});
