/**
 * Tracked changes nested inside a w:hyperlink are always flattened to an
 * editable Hyperlink object, even under revisionHandling:'preserve'. This is
 * a deliberate, documented exception to 'preserve' (inserted/moved-to content
 * kept, deleted/moved-away content dropped, no in-hyperlink revision markup
 * retained on round-trip). This test pins that behavior so the documented
 * exception cannot silently regress.
 */

import { Document } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { PreservedElement } from '../../src/elements/PreservedElement';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function createDocxWithHyperlink(documentXml: string): Promise<Buffer> {
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
  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>`
  );
  zipHandler.addFile('word/document.xml', documentXml);
  return zipHandler.toBuffer();
}

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink r:id="rId5">
        <w:del w:id="10" w:author="T" w:date="2024-01-01T00:00:00Z">
          <w:r><w:delText>OLD-DELETED</w:delText></w:r>
        </w:del>
        <w:ins w:id="11" w:author="T" w:date="2024-01-01T00:00:00Z">
          <w:r><w:t>NEW-INSERTED</w:t></w:r>
        </w:ins>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;

describe('Hyperlink revisions flattened under preserve (X11)', () => {
  it('flattens to an editable Hyperlink, not a PreservedElement, in preserve mode', async () => {
    const buffer = await createDocxWithHyperlink(DOCUMENT_XML);
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    try {
      const content = doc.getParagraphs()[0]!.getContent();
      const hyperlinks = content.filter((item) => item instanceof Hyperlink);
      const preserved = content.filter((item) => item instanceof PreservedElement);

      expect(hyperlinks.length).toBe(1);
      expect(preserved.length).toBe(0);
      // Inserted text kept, deleted text dropped.
      expect((hyperlinks[0] as Hyperlink).getText()).toBe('NEW-INSERTED');
    } finally {
      doc.dispose();
    }
  });

  it('drops in-hyperlink revision markup on save even in preserve mode', async () => {
    const buffer = await createDocxWithHyperlink(DOCUMENT_XML);
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const out = await doc.toBuffer();
    doc.dispose();

    const outZip = new ZipHandler();
    await outZip.loadFromBuffer(out);
    const outXml = outZip.getFileAsString('word/document.xml') ?? '';

    expect(outXml).toContain('NEW-INSERTED');
    expect(outXml).not.toContain('OLD-DELETED');
    // No deletion markup survives inside the hyperlink.
    expect(outXml).not.toContain('w:delText');
  });
});
