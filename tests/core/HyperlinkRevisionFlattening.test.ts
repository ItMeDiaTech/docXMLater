/**
 * Tests for hyperlink revision flattening during parsing.
 *
 * When a w:hyperlink contains tracked change children (w:ins, w:del,
 * w:moveFrom, w:moveTo), the parser flattens the revisions so the
 * hyperlink is parsed as a normal editable Hyperlink object:
 * - w:ins runs are unwrapped (kept)
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

describe('Hyperlink Revision Flattening', () => {
  it('should flatten w:ins + w:del: keep inserted text, drop deleted text', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:del w:id="10" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:delText>old text</w:delText></w:r>
        </w:del>
        <w:ins w:id="11" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>new text</w:t></w:r>
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
    const preserved = content.filter((item) => item instanceof PreservedElement);

    expect(hyperlinks.length).toBe(1);
    expect(preserved.length).toBe(0);

    // Inserted text kept, deleted text dropped
    expect(hyperlinks[0]!.getText()).toBe('new text');
    expect(hyperlinks[0]!.getUrl()).toBe('https://example.com');

    doc.dispose();
  });

  it('should flatten w:ins only: keep inserted text', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:ins w:id="11" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>inserted only</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxWithHyperlinks(documentXml, [
      { id: 'rId5', target: 'https://example.com/inserted' },
    ]);

    const doc = await Document.loadFromBuffer(buffer, {
      revisionHandling: 'preserve',
    });

    const content = doc.getParagraphs()[0]!.getContent();
    const hyperlinks = content.filter((item) => item instanceof Hyperlink);

    expect(hyperlinks.length).toBe(1);
    expect(hyperlinks[0]!.getText()).toBe('inserted only');
    expect(hyperlinks[0]!.getUrl()).toBe('https://example.com/inserted');

    doc.dispose();
  });

  it('should combine direct w:r and w:ins runs', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:r><w:t>direct </w:t></w:r>
        <w:ins w:id="11" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>inserted</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createDocxWithHyperlinks(documentXml, [
      { id: 'rId5', target: 'https://example.com/combined' },
    ]);

    const doc = await Document.loadFromBuffer(buffer, {
      revisionHandling: 'preserve',
    });

    const content = doc.getParagraphs()[0]!.getContent();
    const hyperlinks = content.filter((item) => item instanceof Hyperlink);

    expect(hyperlinks.length).toBe(1);
    // Both direct run text and inserted run text should be present
    const text = hyperlinks[0]!.getText();
    expect(text).toContain('direct');
    expect(text).toContain('inserted');

    doc.dispose();
  });

  it('should produce an editable Hyperlink that supports setUrl() and setText()', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:del w:id="10" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:delText>old</w:delText></w:r>
        </w:del>
        <w:ins w:id="11" w:author="TestUser" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>original</w:t></w:r>
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

    const hyperlink = hyperlinks[0]!;

    // Modify the hyperlink — this would fail on a PreservedElement
    hyperlink.setUrl('https://updated.example.com');
    hyperlink.setText('updated text');

    expect(hyperlink.getUrl()).toBe('https://updated.example.com');
    expect(hyperlink.getText()).toBe('updated text');

    // Round-trip: save and verify the modifications persist
    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml') || '';

    expect(outputXml).toContain('updated text');
    expect(outputXml).toContain('w:hyperlink');
    // Deleted text should not appear
    expect(outputXml).not.toContain('old');
  });
});
