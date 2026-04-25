/**
 * `<w:bookmarkEnd w:displacedByCustomXml="…"/>` — round-trip coverage.
 *
 * Per ECMA-376 Part 1 §17.13.5 CT_MarkupRange, `w:displacedByCustomXml`
 * is a ST_DisplacedByCustomXml attribute ("next" | "prev") valid on
 * BOTH `<w:bookmarkStart>` (CT_Bookmark) and `<w:bookmarkEnd>`
 * (CT_MarkupRange). Word emits it when a bookmark marker had to be
 * displaced because a custom-XML node boundary fell at the same
 * character position — the attribute disambiguates which side of the
 * boundary the marker semantically belongs to.
 *
 * The Bookmark model already preserves the attribute on toStartXML/
 * toEndXML. The parser, however, read it only from `<w:bookmarkStart>`
 * (three parse sites in the codebase); every bookmarkEnd parser
 * silently dropped it on load. Round-trip loss: the displacement
 * disambiguator on the end marker reverted to "absent" on save, which
 * can change where Word positions the range on reload.
 *
 * Iteration 101 extends the three bookmarkEnd parse sites:
 *   - parseBookmarkEnd (XML-string path used for revisions)
 *   - parseHyperlinkFromObject's embedded bookmarkEnd handling
 *   - extractBookmarkEndsFromContent (body-level scan)
 * All three now preserve the "next" | "prev" marker.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndResaveDocXml(xml: string): Promise<string> {
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
  zipHandler.addFile('word/document.xml', xml);
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const content = zip.getFile('word/document.xml')?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

describe('<w:bookmarkEnd> w:displacedByCustomXml round-trip', () => {
  it('preserves w:displacedByCustomXml="next" on bookmarkEnd', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="1" w:name="b1"/>
      <w:r><w:t>x</w:t></w:r>
      <w:bookmarkEnd w:id="1" w:displacedByCustomXml="next"/>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    expect(out).toMatch(/<w:bookmarkEnd[^/]*w:displacedByCustomXml="next"/);
  });

  it('preserves w:displacedByCustomXml="prev" on bookmarkEnd', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="2" w:name="b2"/>
      <w:r><w:t>x</w:t></w:r>
      <w:bookmarkEnd w:id="2" w:displacedByCustomXml="prev"/>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    expect(out).toMatch(/<w:bookmarkEnd[^/]*w:displacedByCustomXml="prev"/);
  });

  it('still emits bookmarkEnd without attribute when absent (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="3" w:name="b3"/>
      <w:r><w:t>x</w:t></w:r>
      <w:bookmarkEnd w:id="3"/>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const endTag = out.match(/<w:bookmarkEnd[^>]*\/>/)?.[0] ?? '';
    expect(endTag).toMatch(/w:id="3"/);
    expect(endTag).not.toMatch(/w:displacedByCustomXml/);
  });

  it('ignores malformed w:displacedByCustomXml values', async () => {
    // Only "next" / "prev" are valid per ST_DisplacedByCustomXml.
    // Unknown values (e.g., "invalid") must fall through to undefined
    // rather than propagating a spec-invalid attribute on save.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="4" w:name="b4"/>
      <w:r><w:t>x</w:t></w:r>
      <w:bookmarkEnd w:id="4" w:displacedByCustomXml="invalid"/>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const endTag = out.match(/<w:bookmarkEnd[^>]*\/>/)?.[0] ?? '';
    expect(endTag).not.toMatch(/w:displacedByCustomXml="invalid"/);
  });
});
