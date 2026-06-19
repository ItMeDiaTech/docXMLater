/**
 * Run text was XML-unescaped twice: XMLParser.parseElementToObject already
 * unescapes entities while building #text, and extractTextValue unescaped the
 * same string again. Text that literally contains an entity sequence (e.g. a
 * document quoting escaped HTML) was therefore corrupted on a pure load/save
 * round-trip: authored "&lt;tag&gt;" (stored in XML as "&amp;lt;tag&amp;gt;")
 * collapsed to "<tag>". extractTextValue no longer double-unescapes.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndResave(documentXml: string): Promise<{ text: string; xml: string }> {
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
  zipHandler.addFile('word/document.xml', documentXml);
  const buffer = await zipHandler.toBuffer();

  const doc = await Document.loadFromBuffer(buffer);
  const text = doc
    .getAllParagraphs()
    .map((p) => p.getText())
    .join('');
  const out = await doc.toBuffer();
  doc.dispose();

  const outZip = new ZipHandler();
  await outZip.loadFromBuffer(out);
  return { text, xml: outZip.getFileAsString('word/document.xml') ?? '' };
}

describe('Run text is not double-unescaped (C39)', () => {
  it('preserves literal entity sequences in the in-model text', async () => {
    // Source text the user authored: code: &lt;tag&gt; and AT&amp;T
    // Stored in XML with one more layer of escaping.
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">code: &amp;lt;tag&amp;gt; and AT&amp;amp;T</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const { text } = await loadAndResave(documentXml);

    // After a SINGLE unescape the literal text must still read as the
    // authored entity sequence — NOT collapse to "code: <tag> and AT&T".
    expect(text).toBe('code: &lt;tag&gt; and AT&amp;T');
    expect(text).not.toBe('code: <tag> and AT&T');
  });

  it('re-emits the doubly-escaped form on save (round-trip stable)', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">code: &amp;lt;tag&amp;gt; and AT&amp;amp;T</w:t></w:r></w:p>
  </w:body>
</w:document>`;

    const { xml } = await loadAndResave(documentXml);

    // The saved text node must keep the escaped entity sequences intact.
    expect(xml).toContain('&amp;lt;tag&amp;gt;');
    expect(xml).toContain('AT&amp;amp;T');
  });
});
