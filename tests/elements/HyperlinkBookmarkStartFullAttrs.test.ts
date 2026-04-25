/**
 * `<w:hyperlink>` wrapping a `<w:bookmarkStart>` — the embedded
 * bookmarkStart parser was missing three CT_Bookmark attributes.
 *
 * Per ECMA-376 Part 1 §17.16.5 CT_Bookmark, `<w:bookmarkStart>` carries:
 *   w:id (required), w:name (required), w:colFirst, w:colLast,
 *   w:displacedByCustomXml
 *
 * The top-level parseBookmarkStart XML-string parser reads all five. The
 * embedded object-form parser inside `parseHyperlinkFromObject` read only
 * w:id and w:name — so whenever a hyperlink wrapped a bookmark (common
 * pattern: a table-column-scoped cross-reference link), the colFirst /
 * colLast / displacedByCustomXml markers were silently dropped on load.
 *
 * Iteration 103 mirrors the XML-string parser, preserving all five
 * attributes on the hyperlink-embedded path.
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

describe('<w:hyperlink>-embedded <w:bookmarkStart> CT_Bookmark full round-trip', () => {
  it('preserves w:colFirst and w:colLast on a bookmarkStart inside a hyperlink', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink w:anchor="target">
        <w:bookmarkStart w:id="10" w:name="colBM" w:colFirst="0" w:colLast="2"/>
        <w:r><w:t>link</w:t></w:r>
        <w:bookmarkEnd w:id="10"/>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const startTag = out.match(/<w:bookmarkStart[^/]*\/>/)?.[0] ?? '';
    expect(startTag).toMatch(/w:name="colBM"/);
    expect(startTag).toMatch(/w:colFirst="0"/);
    expect(startTag).toMatch(/w:colLast="2"/);
  });

  it('preserves w:displacedByCustomXml on a bookmarkStart inside a hyperlink', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink w:anchor="target">
        <w:bookmarkStart w:id="11" w:name="dxBM" w:displacedByCustomXml="next"/>
        <w:r><w:t>link</w:t></w:r>
        <w:bookmarkEnd w:id="11"/>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const startTag = out.match(/<w:bookmarkStart[^/]*\/>/)?.[0] ?? '';
    expect(startTag).toMatch(/w:name="dxBM"/);
    expect(startTag).toMatch(/w:displacedByCustomXml="next"/);
  });

  it('still parses a plain bookmarkStart inside a hyperlink without extras (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:hyperlink w:anchor="target">
        <w:bookmarkStart w:id="12" w:name="plain"/>
        <w:r><w:t>link</w:t></w:r>
        <w:bookmarkEnd w:id="12"/>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const startTag = out.match(/<w:bookmarkStart[^/]*\/>/)?.[0] ?? '';
    expect(startTag).toMatch(/w:name="plain"/);
    expect(startTag).not.toMatch(/w:colFirst/);
    expect(startTag).not.toMatch(/w:colLast/);
    expect(startTag).not.toMatch(/w:displacedByCustomXml/);
  });
});
