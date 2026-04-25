/**
 * `<w:sectPrChange>` previous-`<w:cols>` — full CT_Columns round-trip.
 *
 * Per ECMA-376 §17.6.4 CT_Columns has four attributes (num, space,
 * equalWidth, sep) and a sequence of `<w:col w:w="..." w:space="..."/>`
 * children for non-equal-width layouts.
 *
 * The sectPrChange EMITTER already handles `prev.columns.equalWidth`,
 * `prev.columns.separator`, `prev.columns.columnWidths`, and
 * `prev.columns.columnSpaces` — but the PARSER previously only read
 * `w:num` and `w:space`. So any tracked change to a columns layout —
 * e.g. changing the separator from off to on, or switching from equal
 * widths to custom per-column widths, with track-changes enabled — lost
 * the "previous state" of equalWidth / sep / columnWidths / columnSpaces
 * on every load→save round-trip.
 *
 * This iteration closes the parser gap.
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

describe('<w:sectPrChange> previous <w:cols> full round-trip', () => {
  it('preserves equalWidth and sep attributes in sectPrChange history', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:sectPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:cols w:num="3" w:space="720" w:equalWidth="0" w:sep="1"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:cols[^>]*w:num="3"/);
    expect(changeBlock).toMatch(/<w:cols[^>]*w:equalWidth="0"/);
    expect(changeBlock).toMatch(/<w:cols[^>]*w:sep="1"/);
  });

  it('preserves per-column widths (<w:col>) in sectPrChange history', async () => {
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
          <w:cols w:num="2" w:equalWidth="0">
            <w:col w:w="3600" w:space="720"/>
            <w:col w:w="5040"/>
          </w:cols>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    // Both column children must round-trip with their widths.
    expect(changeBlock).toMatch(/<w:col[^>]*w:w="3600"/);
    expect(changeBlock).toMatch(/<w:col[^>]*w:w="5040"/);
    // The column-level space should survive too.
    expect(changeBlock).toMatch(/<w:col[^>]*w:space="720"/);
  });
});
