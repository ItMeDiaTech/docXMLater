/**
 * Tests for tracked changes inside hyperlinks being preserved during round-trip.
 *
 * Issue: w:del/w:ins inside w:hyperlink was lost during parsing because
 * parseHyperlinkFromObject() only processes direct w:r children. On save,
 * the revision markup disappeared, causing deleted text to appear as regular
 * visible text.
 *
 * Fix: Detect revision children inside hyperlink objects during parsing and
 * preserve the entire hyperlink as a PreservedElement (raw XML passthrough).
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
    it('should preserve tracked changes inside hyperlinks during round-trip', async () => {
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

      // The hyperlink should be preserved as a PreservedElement (not parsed as Hyperlink)
      const paras = doc.getParagraphs();
      expect(paras.length).toBeGreaterThanOrEqual(1);
      const content = paras[0]!.getContent();
      const preserved = content.filter((item) => item instanceof PreservedElement);
      expect(preserved.length).toBe(1);
      expect(preserved[0]!.getElementType()).toBe('w:hyperlink');

      // Round-trip: save and verify output XML
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // Must contain the revision elements inside the hyperlink
      expect(outputXml).toContain('w:del');
      expect(outputXml).toContain('w:ins');
      expect(outputXml).toContain('w:delText');
      expect(outputXml).toContain('old link text');
      expect(outputXml).toContain('new link text');
      expect(outputXml).toContain('w:hyperlink');
    });

    it('should preserve w:moveFrom and w:moveTo inside hyperlinks', async () => {
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
      const preserved = content.filter((item) => item instanceof PreservedElement);
      expect(preserved.length).toBe(1);
      expect(preserved[0]!.getRawXml()).toContain('w:moveFrom');

      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';
      expect(outputXml).toContain('w:moveFrom');
      expect(outputXml).toContain('moved text');
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

      // First hyperlink (normal) should be a Hyperlink object
      expect(hyperlinks.length).toBe(1);
      // Second hyperlink (with revisions) should be a PreservedElement
      expect(preserved.length).toBe(1);
      expect(preserved[0]!.getElementType()).toBe('w:hyperlink');

      // Round-trip
      const outputBuffer = await doc.toBuffer();
      doc.dispose();

      const outputZip = new ZipHandler();
      await outputZip.loadFromBuffer(outputBuffer);
      const outputXml = outputZip.getFileAsString('word/document.xml') || '';

      // Both hyperlinks should be present
      expect(outputXml).toContain('normal link');
      expect(outputXml).toContain('w:del');
      expect(outputXml).toContain('w:ins');
      expect(outputXml).toContain('w:delText');
    });

    it('should preserve hyperlink attributes (r:id) in the raw XML', async () => {
      const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5" w:tooltip="Example tooltip">
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
      const preserved = content.filter(
        (item) => item instanceof PreservedElement
      ) as PreservedElement[];
      expect(preserved.length).toBe(1);

      // Raw XML should contain the relationship ID and tooltip
      const rawXml = preserved[0]!.getRawXml();
      expect(rawXml).toContain('r:id="rId5"');
      expect(rawXml).toContain('w:tooltip="Example tooltip"');

      doc.dispose();
    });
  });
});
