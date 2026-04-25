/**
 * Run character border `<w:bdr>` — full CT_Border attribute fidelity.
 *
 * Per ECMA-376 §17.3.2.5, the character border is CT_Border (§17.18.2)
 * so it carries nine attributes: val / sz / space / color / themeColor /
 * themeTint / themeShade / shadow / frame.
 *
 * Iteration 79 extended the three emission paths to include the full set.
 * This iteration extends the three PARSER paths so themed character
 * borders round-trip without loss:
 *   1. Main run parser (object format, parseRunFromObject).
 *   2. rPrChange previous-rPr parser (object format).
 *   3. Style-level rPr parser (XML-string format).
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithRunBdr(bdrAttrs: string): Promise<Buffer> {
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
    <w:p>
      <w:r>
        <w:rPr><w:bdr ${bdrAttrs}/></w:rPr>
        <w:t>bordered</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('Run character border (<w:bdr>) full CT_Border round-trip', () => {
  it('round-trips themeColor / themeTint / themeShade', async () => {
    const buffer = await makeDocxWithRunBdr(
      'w:val="single" w:sz="4" w:color="auto" w:themeColor="accent1" w:themeTint="99" w:themeShade="80"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:bdr[^>]*w:themeColor="accent1"/);
    expect(xml).toMatch(/<w:bdr[^>]*w:themeTint="99"/);
    expect(xml).toMatch(/<w:bdr[^>]*w:themeShade="80"/);
  });

  it('round-trips w:shadow="1" on character border', async () => {
    const buffer = await makeDocxWithRunBdr(
      'w:val="single" w:sz="4" w:color="000000" w:shadow="1"'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:bdr[^>]*w:shadow="1"/);
  });

  it('round-trips w:frame="1" on character border', async () => {
    const buffer = await makeDocxWithRunBdr('w:val="single" w:sz="4" w:color="000000" w:frame="1"');
    const doc = await Document.loadFromBuffer(buffer);
    const out = await doc.toBuffer();
    doc.dispose();

    const zip = new ZipHandler();
    await zip.loadFromBuffer(out);
    const docFile = zip.getFile('word/document.xml');
    const content = docFile?.content;
    const xml = content instanceof Buffer ? content.toString('utf8') : String(content);

    expect(xml).toMatch(/<w:bdr[^>]*w:frame="1"/);
  });
});
