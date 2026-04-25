/**
 * `<w:sectPrChange>` — pgMar `w:gutter` and pgNumType chapStyle/chapSep
 * round-trip.
 *
 * Per ECMA-376:
 *   - §17.6.11 CT_PageMar has seven required attributes — top, right,
 *     bottom, left, header, footer, **gutter**. The sectPrChange parser
 *     previously dropped `w:gutter` entirely.
 *   - §17.6.12 CT_PageNumber has four attributes — fmt, start,
 *     **chapStyle**, **chapSep**. The sectPrChange parser previously
 *     dropped chapStyle and chapSep.
 *
 * Both pairs are preserved on the main sectPr parser, so the
 * sectPrChange gap produced asymmetry: a user changing a book-binding
 * gutter or a chapter-numbering scheme with track-changes enabled
 * lost the "previous" state on round-trip.
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

  const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
  const out = await doc.toBuffer();
  doc.dispose();

  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const docFile = zip.getFile('word/document.xml');
  const content = docFile?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

describe('<w:sectPrChange> pgMar gutter + pgNumType chapStyle/chapSep round-trip', () => {
  it('preserves w:gutter on sectPrChange pgMar', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="720"/>
      <w:sectPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:gutter="1440"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:pgMar[^>]*w:gutter="1440"/);
  });

  it('preserves chapStyle and chapSep on sectPrChange pgNumType', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:sectPrChange w:id="2" w:author="Tester" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:pgNumType w:fmt="decimal" w:start="1" w:chapStyle="2" w:chapSep="emDash"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:pgNumType[^>]*w:chapStyle="2"/);
    expect(changeBlock).toMatch(/<w:pgNumType[^>]*w:chapSep="emDash"/);
  });
});
