/**
 * parseRevisionFromXml must keep interleaved runs and hyperlinks in document
 * order. The previous two-pass extraction (all standalone w:r first, then all
 * w:hyperlink) appended every hyperlink after the last run, so a tracked
 * insertion like <w:ins><w:r>alpha </w:r><w:hyperlink>LINK</w:hyperlink>
 * <w:r> omega</w:r></w:ins> round-tripped under revisionHandling:'preserve'
 * with the text reading "alpha omegaLINK".
 */
import { Document } from '../../src/core/Document';
import { Hyperlink } from '../../src/elements/Hyperlink';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithInterleavedRevision(): Promise<Buffer> {
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
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:ins w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z"><w:r><w:t xml:space="preserve">alpha </w:t></w:r><w:hyperlink r:id="rId5"><w:r><w:t>LINK</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> omega</w:t></w:r></w:ins>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('Revision parsing keeps interleaved run/hyperlink document order', () => {
  it('parses run, hyperlink, run in original order under preserve mode', async () => {
    const buffer = await makeDocxWithInterleavedRevision();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    try {
      const paragraphs = doc.getParagraphs();
      const revisions = paragraphs.flatMap((p) => p.getRevisions());
      expect(revisions).toHaveLength(1);

      const content = revisions[0]!.getContent();
      expect(content).toHaveLength(3);
      expect(content[0]).toBeInstanceOf(Run);
      expect((content[0] as Run).getText()).toBe('alpha ');
      expect(content[1]).toBeInstanceOf(Hyperlink);
      expect(content[2]).toBeInstanceOf(Run);
      expect((content[2] as Run).getText()).toBe(' omega');
    } finally {
      doc.dispose();
    }
  });

  it('round-trips revision text in document order (alpha, LINK, omega)', async () => {
    const buffer = await makeDocxWithInterleavedRevision();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    try {
      const saved = await doc.toBuffer();
      const outZip = new ZipHandler();
      await outZip.loadFromBuffer(saved);
      const savedXml = outZip.getFileAsString('word/document.xml') ?? '';

      const alphaIdx = savedXml.indexOf('alpha ');
      const linkIdx = savedXml.indexOf('LINK');
      const omegaIdx = savedXml.indexOf(' omega');
      expect(alphaIdx).toBeGreaterThan(-1);
      expect(linkIdx).toBeGreaterThan(alphaIdx);
      expect(omegaIdx).toBeGreaterThan(linkIdx);
    } finally {
      doc.dispose();
    }
  });

  it('keeps order when the hyperlink leads and trails runs', async () => {
    const zipHandler = new ZipHandler();
    const base = await makeDocxWithInterleavedRevision();
    await zipHandler.loadFromBuffer(base);
    zipHandler.addFile(
      'word/document.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:ins w:id="2" w:author="A" w:date="2026-01-15T10:00:00Z"><w:hyperlink r:id="rId5"><w:r><w:t>FIRST</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> middle </w:t></w:r><w:hyperlink r:id="rId5"><w:r><w:t>LAST</w:t></w:r></w:hyperlink></w:ins>
    </w:p>
  </w:body>
</w:document>`
    );
    const doc = await Document.loadFromBuffer(await zipHandler.toBuffer(), {
      revisionHandling: 'preserve',
    });
    try {
      const revisions = doc.getParagraphs().flatMap((p) => p.getRevisions());
      expect(revisions).toHaveLength(1);
      const content = revisions[0]!.getContent();
      expect(content).toHaveLength(3);
      expect(content[0]).toBeInstanceOf(Hyperlink);
      expect((content[1] as Run).getText()).toBe(' middle ');
      expect(content[2]).toBeInstanceOf(Hyperlink);
    } finally {
      doc.dispose();
    }
  });
});
