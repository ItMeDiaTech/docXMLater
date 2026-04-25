/**
 * `<w:sectPrChange>` previous-`<w:sectPr>` round-trip for footnotePr /
 * endnotePr / paperSrc.
 *
 * Per ECMA-376 §17.13.5.32 CT_SectPrChange carries a child `<w:sectPr>`
 * (CT_SectPrBase) that can contain all the same CT_SectPrBase children
 * as the direct sectPr — including `<w:footnotePr>` (§17.11.9),
 * `<w:endnotePr>` (§17.11.5), and `<w:paperSrc>` (§17.6.12).
 *
 * The sectPrChange emitter already supports `prev.footnotePr`,
 * `prev.endnotePr`, and `prev.paperSource` — but the PARSER never
 * extracted them. Consequence: a tracked change to any of these
 * three section-level properties lost its "previous state" on every
 * load→save round-trip, so reviewers couldn't see what was changed.
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

describe('<w:sectPrChange> previous <w:footnotePr> / <w:endnotePr> / <w:paperSrc>', () => {
  it('preserves footnotePr in sectPrChange history', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:sectPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:footnotePr>
            <w:pos w:val="pageBottom"/>
            <w:numFmt w:val="lowerRoman"/>
            <w:numStart w:val="3"/>
            <w:numRestart w:val="eachSect"/>
          </w:footnotePr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:footnotePr>/);
    expect(changeBlock).toMatch(/<w:pos\s+w:val="pageBottom"/);
    expect(changeBlock).toMatch(/<w:numFmt\s+w:val="lowerRoman"/);
    expect(changeBlock).toMatch(/<w:numStart\s+w:val="3"/);
    expect(changeBlock).toMatch(/<w:numRestart\s+w:val="eachSect"/);
  });

  it('preserves endnotePr in sectPrChange history', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:sectPrChange w:id="2" w:author="Tester" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:endnotePr>
            <w:pos w:val="docEnd"/>
            <w:numFmt w:val="upperRoman"/>
          </w:endnotePr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:endnotePr>/);
    expect(changeBlock).toMatch(/<w:pos\s+w:val="docEnd"/);
    expect(changeBlock).toMatch(/<w:numFmt\s+w:val="upperRoman"/);
  });

  it('preserves paperSrc in sectPrChange history', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>content</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:sectPrChange w:id="3" w:author="Tester" w:date="2026-01-01T00:00:00Z">
        <w:sectPr>
          <w:pgSz w:w="12240" w:h="15840"/>
          <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
          <w:paperSrc w:first="7" w:other="3"/>
        </w:sectPr>
      </w:sectPrChange>
    </w:sectPr>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:sectPrChange[\s\S]*?<\/w:sectPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:paperSrc[^>]*w:first="7"/);
    expect(changeBlock).toMatch(/<w:paperSrc[^>]*w:other="3"/);
  });
});
