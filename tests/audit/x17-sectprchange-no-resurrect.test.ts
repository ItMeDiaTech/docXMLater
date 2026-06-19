/**
 * parseSectionProperties ran every per-property extractElements over the whole
 * sectPr string, which includes the nested previous <w:sectPr> inside
 * <w:sectPrChange>. When a tracked change REMOVED a property (absent from the
 * live sectPr, present only in the previous one), the flat scan matched the
 * nested previous value and applied it as a live section property —
 * resurrecting a tracked-removed property (e.g. titlePg/bidi). The change
 * subtree is now stripped before reading live properties.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndResaveDocXml(
  xml: string,
  options?: { revisionHandling?: 'accept' | 'strip' | 'preserve'; acceptRevisions?: boolean }
): Promise<string> {
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
  const doc = await Document.loadFromBuffer(buffer, options ?? { revisionHandling: 'preserve' });
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  return zip.getFileAsString('word/document.xml') ?? '';
}

// Live sectPr is everything before <w:sectPrChange>; the change block is the
// nested previous state. We assert on the live portion only.
function liveSectPr(xml: string): string {
  const sectPr = xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0] ?? '';
  const changeStart = sectPr.indexOf('<w:sectPrChange');
  return changeStart === -1 ? sectPr : sectPr.slice(0, changeStart);
}

// Source: live sectPr has NO titlePg/bidi; the previous sectPr inside
// sectPrChange has both. They must stay in the change history only.
const XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:sectPrChange w:id="1" w:author="T" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:titlePg/>
          <w:bidi/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;

describe('sectPrChange previous properties are not resurrected as live (X17)', () => {
  it('does not add titlePg/bidi to the live sectPr under preserve', async () => {
    const out = await loadAndResaveDocXml(XML, { revisionHandling: 'preserve' });
    const live = liveSectPr(out);

    expect(live).not.toContain('<w:titlePg');
    expect(live).not.toContain('<w:bidi');
    // The previous-state markers still live in the change block.
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toContain('<w:titlePg');
    expect(changeBlock).toContain('<w:bidi');
  });

  it('does not resurrect titlePg/bidi under acceptRevisions:true', async () => {
    const out = await loadAndResaveDocXml(XML, { acceptRevisions: true });
    const live = liveSectPr(out);

    expect(live).not.toContain('<w:titlePg');
    expect(live).not.toContain('<w:bidi');
  });
});
