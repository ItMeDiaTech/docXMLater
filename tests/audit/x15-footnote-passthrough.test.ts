/**
 * Footnote/endnote round-trip fidelity:
 * - unmodified loaded notes must pass through byte-identically (dirty flag
 *   gates regeneration, not the registered-note count)
 * - regeneration after a real edit must keep paragraph styles (pStyle
 *   FootnoteText per ECMA-376 §17.7.4) and w:hyperlink wrappers with their
 *   part-scoped r:id relationships
 */

import { Document } from '../../src/core/Document';

const JSZip = require('jszip');

const FOOTNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> See </w:t></w:r><w:hyperlink r:id="rId100"><w:r><w:t>example</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> for details</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

const FOOTNOTES_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>`;

const ENDNOTES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  <w:endnote w:id="1"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> See </w:t></w:r><w:hyperlink r:id="rId200"><w:r><w:t>endnote-link</w:t></w:r></w:hyperlink></w:p></w:endnote>
</w:endnotes>`;

const ENDNOTES_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId200" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/" TargetMode="External"/>
</Relationships>`;

async function buildFixture(): Promise<Buffer> {
  const doc = Document.create();
  doc.createParagraph('Body text');
  const buf = await doc.toBuffer();
  doc.dispose();

  const zip = await JSZip.loadAsync(buf);
  zip.file('word/footnotes.xml', FOOTNOTES_XML);
  zip.file('word/_rels/footnotes.xml.rels', FOOTNOTES_RELS);
  zip.file('word/endnotes.xml', ENDNOTES_XML);
  zip.file('word/_rels/endnotes.xml.rels', ENDNOTES_RELS);
  return zip.generateAsync({ type: 'nodebuffer' });
}

describe('X15: footnotes/endnotes passthrough and lossless regeneration', () => {
  it('preserves footnotes.xml byte-identically on plain load->save', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      const outBuf = await doc.toBuffer();
      const outZip = await JSZip.loadAsync(outBuf);
      const outFootnotes = await outZip.file('word/footnotes.xml')?.async('string');

      expect(outFootnotes).toBe(FOOTNOTES_XML);
      expect(outFootnotes).toContain('<w:pStyle w:val="FootnoteText"/>');
      expect(outFootnotes).toContain('<w:hyperlink r:id="rId100">');
    } finally {
      doc.dispose();
    }
  });

  it('preserves endnotes.xml byte-identically on plain load->save', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      const outBuf = await doc.toBuffer();
      const outZip = await JSZip.loadAsync(outBuf);
      const outEndnotes = await outZip.file('word/endnotes.xml')?.async('string');

      expect(outEndnotes).toBe(ENDNOTES_XML);
      expect(outEndnotes).toContain('<w:pStyle w:val="EndnoteText"/>');
      expect(outEndnotes).toContain('<w:hyperlink r:id="rId200">');
    } finally {
      doc.dispose();
    }
  });

  it('keeps pStyle and hyperlink wrappers when footnotes are regenerated after an edit', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      doc.createFootnote('Brand new note');
      const outBuf = await doc.toBuffer();
      const outZip = await JSZip.loadAsync(outBuf);
      const outFootnotes = await outZip.file('word/footnotes.xml')?.async('string');
      const outRels = await outZip.file('word/_rels/footnotes.xml.rels')?.async('string');

      expect(outFootnotes).toContain('Brand new note');
      expect(outFootnotes).toContain('w:pStyle w:val="FootnoteText"');
      expect(outFootnotes).toContain('<w:hyperlink');
      expect(outFootnotes).toContain('r:id="rId100"');
      expect(outFootnotes).toContain('example');
      // Hyperlink relationship must survive in the part-scoped rels file
      expect(outRels).toContain('rId100');
      expect(outRels).toContain('https://example.com/');
    } finally {
      doc.dispose();
    }
  });

  it('keeps pStyle and hyperlink wrappers when endnotes are regenerated after an edit', async () => {
    const fixture = await buildFixture();
    const doc = await Document.loadFromBuffer(fixture);
    try {
      doc.createEndnote('Brand new endnote');
      const outBuf = await doc.toBuffer();
      const outZip = await JSZip.loadAsync(outBuf);
      const outEndnotes = await outZip.file('word/endnotes.xml')?.async('string');
      const outRels = await outZip.file('word/_rels/endnotes.xml.rels')?.async('string');

      expect(outEndnotes).toContain('Brand new endnote');
      expect(outEndnotes).toContain('w:pStyle w:val="EndnoteText"');
      expect(outEndnotes).toContain('<w:hyperlink');
      expect(outEndnotes).toContain('r:id="rId200"');
      expect(outRels).toContain('rId200');
      expect(outRels).toContain('https://example.org/');
    } finally {
      doc.dispose();
    }
  });
});
