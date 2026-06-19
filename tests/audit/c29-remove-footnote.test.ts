/**
 * Document.removeFootnote/removeEndnote:
 * - the definition is removed from the saved part (even when it was the last
 *   one — the dirty flag must defeat the original-XML passthrough)
 * - the matching w:footnoteReference/w:endnoteReference run is stripped from
 *   the body so every remaining reference resolves to a definition per
 *   ECMA-376 §17.11.14 / §17.11.3
 */

import { Document } from '../../src/core/Document';

const JSZip = require('jszip');

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">Text before</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> middle </w:t></w:r><w:r><w:footnoteReference w:id="3"/></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">Endnote anchor</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;

const FOOTNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>Doomed footnote</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="3"><w:p><w:r><w:t>Surviving footnote</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

const ENDNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:id="2"><w:p><w:r><w:t>Doomed endnote</w:t></w:r></w:p></w:endnote>
</w:endnotes>`;

async function buildFixture(): Promise<Buffer> {
  const doc = Document.create();
  doc.createParagraph('Placeholder');
  const buf = await doc.toBuffer();
  doc.dispose();

  const zip = await JSZip.loadAsync(buf);
  zip.file('word/document.xml', DOCUMENT_XML);
  zip.file('word/footnotes.xml', FOOTNOTES_XML);
  zip.file('word/endnotes.xml', ENDNOTES_XML);
  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('C29: Document.removeFootnote/removeEndnote', () => {
  it('removes the definition and strips the body footnoteReference', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      expect(doc.removeFootnote(2)).toBe(true);

      const outBuf = await doc.toBuffer();
      const outZip = await JSZip.loadAsync(outBuf);
      const outFootnotes = await outZip.file('word/footnotes.xml')?.async('string');
      const outDocument = await outZip.file('word/document.xml')?.async('string');

      // Definition gone, sibling definition intact
      expect(outFootnotes).not.toContain('Doomed footnote');
      expect(outFootnotes).toContain('Surviving footnote');

      // Body reference stripped, sibling reference and text intact
      expect(outDocument).not.toContain('<w:footnoteReference w:id="2"/>');
      expect(outDocument).toContain('<w:footnoteReference w:id="3"/>');
      expect(outDocument).toContain('Text before');
    } finally {
      doc.dispose();
    }
  });

  it('removing the last footnote is not undone by the original-XML passthrough', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      expect(doc.removeFootnote(2)).toBe(true);
      expect(doc.removeFootnote(3)).toBe(true);

      const outBuf = await doc.toBuffer();
      const outZip = await JSZip.loadAsync(outBuf);
      const outFootnotes = await outZip.file('word/footnotes.xml')?.async('string');
      const outDocument = await outZip.file('word/document.xml')?.async('string');

      expect(outFootnotes).toBeDefined();
      expect(outFootnotes).not.toContain('Doomed footnote');
      expect(outFootnotes).not.toContain('Surviving footnote');
      // Separators must survive (footnotes.xml regenerated, not deleted)
      expect(outFootnotes).toContain('<w:separator/>');
      expect(outDocument).not.toContain('w:footnoteReference');
    } finally {
      doc.dispose();
    }
  });

  it('removes the definition and strips the body endnoteReference', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      expect(doc.removeEndnote(2)).toBe(true);

      const outBuf = await doc.toBuffer();
      const outZip = await JSZip.loadAsync(outBuf);
      const outEndnotes = await outZip.file('word/endnotes.xml')?.async('string');
      const outDocument = await outZip.file('word/document.xml')?.async('string');

      expect(outEndnotes).toBeDefined();
      expect(outEndnotes).not.toContain('Doomed endnote');
      expect(outDocument).not.toContain('w:endnoteReference');
      expect(outDocument).toContain('Endnote anchor');
    } finally {
      doc.dispose();
    }
  });

  it('returns false for unknown ids and protected special notes', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      expect(doc.removeFootnote(999)).toBe(false);
      expect(doc.removeFootnote(-1)).toBe(false);
      expect(doc.removeEndnote(999)).toBe(false);
      expect(doc.removeEndnote(-1)).toBe(false);
      // Untouched ids must survive a rejected removal
      expect(doc.getFootnoteManager().hasFootnote(2)).toBe(true);
      expect(doc.getEndnoteManager().hasEndnote(2)).toBe(true);
    } finally {
      doc.dispose();
    }
  });
});
