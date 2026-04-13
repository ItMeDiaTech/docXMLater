/**
 * Tests for tracked changes inside hyperlinks being flattened during parsing.
 *
 * Revision children (w:ins, w:del, w:moveFrom, w:moveTo) inside w:hyperlink
 * are flattened so the hyperlink is parsed as a normal editable Hyperlink object:
 * - w:ins runs are unwrapped (kept as direct runs)
 * - w:del runs are dropped (deleted content discarded)
 * - w:moveFrom runs are dropped (moved-away content)
 * - w:moveTo runs are unwrapped (move destination kept)
 */

import { Document } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { PreservedElement } from '../../src/elements/PreservedElement';
import { ZipHandler } from '../../src/zip/ZipHandler';

/**
 * Helper to create a minimal DOCX buffer with hyperlink relationships
 */
async function createDocxWithHyperlinks(
  documentXml: string,
  rels: { id: string; target: string }[] = []
): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  const relEntries = rels
    .map(
      (r) =>
        `<Relationship Id="${r.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${r.target}" TargetMode="External"/>`
    )
    .join('\n  ');

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${relEntries}
</Relationships>`
  );

  zipHandler.addFile('word/document.xml', documentXml);

  return await zipHandler.toBuffer();
}

describe('Hyperlink Revision Preservation', () => {
  describe('w:del and w:ins inside w:hyperlink', () => {
    it('should flatten tracked changes inside hyperlinks and parse as Hyperlink', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:del w:id="10" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:delText>old link text</w:delText></w:r>
        </w:del>
        <w:ins w:id="11" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>new link text</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithHyperlinks(documentXml, [
        { id: 'rId5', target: 'https://example.com' },
      ]);

      const doc = await Document.loadFromBuffer(buffer, {
        revisionHandling: 'preserve',
      });

      // The hyperlink should be parsed as a Hyperlink (not PreservedElement)
      const paras = doc.getParagraphs();
      expect(paras.length).toBeGreaterThanOrEqual(1);
      const content = paras[0]!.getContent();
      const hyperlinks = content.filter((item) => item instanceof Hyperlink);
      const preserved = content.filter((item) => item instanceof PreservedElement);
      expect(hyperlinks.length).toBe(1);
      expect(preserved.length).toBe(0);

      // Should contain the inserted text (w:ins unwrapped), deleted text dropped
      expect(hyperlinks[0]!.getText()).toBe('new link text');

      doc.dispose();
    });

    it('should handle w:moveFrom inside hyperlinks by dropping moved-away content', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:moveFrom w:id="20" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>moved text</w:t></w:r>
        </w:moveFrom>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithHyperlinks(documentXml, [
        { id: 'rId5', target: 'https://example.com' },
      ]);

      const doc = await Document.loadFromBuffer(buffer, {
        revisionHandling: 'preserve',
      });

      const content = doc.getParagraphs()[0]!.getContent();
      // moveFrom content is dropped, so the hyperlink has no runs.
      // parseHyperlinkFromObject may not produce a Hyperlink when there are no runs.
      const preserved = content.filter((item) => item instanceof PreservedElement);
      expect(preserved.length).toBe(0);

      doc.dispose();
    });
  });

  describe('normal hyperlinks without revisions', () => {
    it('should still parse normal hyperlinks as Hyperlink objects', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:r><w:t>click here</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithHyperlinks(documentXml, [
        { id: 'rId5', target: 'https://example.com' },
      ]);

      const doc = await Document.loadFromBuffer(buffer, {
        revisionHandling: 'preserve',
      });

      const content = doc.getParagraphs()[0]!.getContent();
      const hyperlinks = content.filter((item) => item instanceof Hyperlink);
      const preserved = content.filter((item) => item instanceof PreservedElement);

      expect(hyperlinks.length).toBe(1);
      expect(preserved.length).toBe(0);

      doc.dispose();
    });
  });

  describe('mixed paragraphs', () => {
    it('should handle both normal and revision-containing hyperlinks in the same paragraph', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:r><w:t>normal link</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t> some text </w:t></w:r>
      <w:hyperlink r:id="rId6">
        <w:del w:id="30" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:delText>old</w:delText></w:r>
        </w:del>
        <w:ins w:id="31" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>new</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithHyperlinks(documentXml, [
        { id: 'rId5', target: 'https://example.com' },
        { id: 'rId6', target: 'https://example.org' },
      ]);

      const doc = await Document.loadFromBuffer(buffer, {
        revisionHandling: 'preserve',
      });

      const content = doc.getParagraphs()[0]!.getContent();
      const hyperlinks = content.filter((item) => item instanceof Hyperlink);
      const preserved = content.filter((item) => item instanceof PreservedElement);

      // Both hyperlinks should be parsed as Hyperlink objects
      expect(hyperlinks.length).toBe(2);
      expect(preserved.length).toBe(0);

      // Second hyperlink should have the inserted text, not the deleted text
      expect(hyperlinks[1]!.getText()).toBe('new');

      doc.dispose();
    });

    it('should preserve hyperlink URL after flattening revisions', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:ins w:id="40" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>inserted text</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

      const buffer = await createDocxWithHyperlinks(documentXml, [
        { id: 'rId5', target: 'https://example.com' },
      ]);

      const doc = await Document.loadFromBuffer(buffer, {
        revisionHandling: 'preserve',
      });

      const content = doc.getParagraphs()[0]!.getContent();
      const hyperlinks = content.filter((item) => item instanceof Hyperlink);
      expect(hyperlinks.length).toBe(1);

      // Should have the correct URL and text
      expect(hyperlinks[0]!.getUrl()).toBe('https://example.com');
      expect(hyperlinks[0]!.getText()).toBe('inserted text');

      doc.dispose();
    });
  });
});
