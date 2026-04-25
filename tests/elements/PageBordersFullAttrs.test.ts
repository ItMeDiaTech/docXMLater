/**
 * `<w:pgBorders>` — full CT_Border attribute set fidelity.
 *
 * CT_TopBorder / CT_BottomBorder (§17.6.2) extend CT_Border (§17.18.2),
 * so page borders carry the same nine attributes as any other border:
 *   w:val (required), w:sz, w:space, w:color, w:themeColor, w:themeTint,
 *   w:themeShade, w:shadow, w:frame.
 * Plus CT_PageBorder adds `w:id` (for art borders).
 *
 * Two bugs existed before this iteration:
 *   1. Parser only stored `shadow`/`frame` when `true` — silently
 *      dropped `w:shadow="0"` / `w:frame="0"`.
 *   2. Parser never read `w:themeTint` / `w:themeShade`.
 *   3. Emitter never emitted `w:themeTint` / `w:themeShade`.
 *   4. Emitter dropped explicit-false shadow/frame via truthy gate.
 *
 * This test covers the full round-trip for all four gaps.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithPgBorders(borderAttrs: string): Promise<Buffer> {
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
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>x</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:pgBorders w:offsetFrom="page">
        <w:top ${borderAttrs}/>
      </w:pgBorders>
    </w:sectPr>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('pgBorders full CT_Border attribute fidelity', () => {
  it('round-trips themeColor / themeTint / themeShade on page border', async () => {
    const buffer = await makeDocxWithPgBorders(
      'w:val="single" w:sz="24" w:color="auto" w:themeColor="accent1" w:themeTint="66" w:themeShade="80"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:pgBorders[^>]*>[\s\S]*<w:top[^>]*w:themeColor="accent1"/);
    expect(xml).toMatch(/<w:top[^>]*w:themeTint="66"/);
    expect(xml).toMatch(/<w:top[^>]*w:themeShade="80"/);
  });

  it('round-trips explicit w:shadow="0" (not dropped)', async () => {
    const buffer = await makeDocxWithPgBorders(
      'w:val="single" w:sz="24" w:color="000000" w:shadow="0"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:top[^>]*w:shadow="0"/);
  });

  it('round-trips w:shadow="1" on page border', async () => {
    const buffer = await makeDocxWithPgBorders(
      'w:val="single" w:sz="24" w:color="000000" w:shadow="1"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:top[^>]*w:shadow="1"/);
  });
});
