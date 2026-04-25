/**
 * `<w:tblPrEx>` table border — full CT_Border attribute round-trip.
 *
 * Per ECMA-376 §17.4.62 CT_TblPrExBase, `<w:tblBorders>` child is CT_TblBorders
 * whose per-side children are CT_Border (§17.18.2) — the full nine-attribute
 * set: val / sz / space / color / themeColor / themeTint / themeShade /
 * shadow / frame.
 *
 * Iteration 79 extended the emitter (`TableRow.buildBordersXML` used for
 * tblPrEx/tblBorders). This iteration extends the PARSER
 * (`parseTableBordersFromObject`) which was previously only reading the
 * four basic attrs — so themed tblPrEx borders were silently stripped on
 * load, leaving the emitter nothing to re-emit.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithTblPrExBorder(borderAttrs: string): Promise<Buffer> {
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
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tblPrEx>
          <w:tblBorders><w:top ${borderAttrs}/></w:tblBorders>
        </w:tblPrEx>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>cell</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>doc</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

async function emittedDocXml(buffer: Buffer): Promise<string> {
  const doc = await Document.loadFromBuffer(buffer);
  const out = await doc.toBuffer();
  doc.dispose();
  const zip = new ZipHandler();
  await zip.loadFromBuffer(out);
  const docFile = zip.getFile('word/document.xml');
  const content = docFile?.content;
  return content instanceof Buffer ? content.toString('utf8') : String(content);
}

describe('<w:tblPrEx> table border — full CT_Border round-trip', () => {
  it('round-trips themeColor / themeTint / themeShade', async () => {
    const buffer = await makeDocxWithTblPrExBorder(
      'w:val="single" w:sz="4" w:color="auto" w:themeColor="accent3" w:themeTint="99" w:themeShade="80"'
    );
    const xml = await emittedDocXml(buffer);
    expect(xml).toMatch(/<w:tblPrEx>[\s\S]*<w:top[^>]*w:themeColor="accent3"/);
    expect(xml).toMatch(/<w:top[^>]*w:themeTint="99"/);
    expect(xml).toMatch(/<w:top[^>]*w:themeShade="80"/);
  });

  it('round-trips w:shadow="1" on tblPrEx border', async () => {
    const buffer = await makeDocxWithTblPrExBorder(
      'w:val="single" w:sz="4" w:color="000000" w:shadow="1"'
    );
    const xml = await emittedDocXml(buffer);
    expect(xml).toMatch(/<w:tblPrEx>[\s\S]*<w:top[^>]*w:shadow="1"/);
  });

  it('round-trips w:frame="1" on tblPrEx border', async () => {
    const buffer = await makeDocxWithTblPrExBorder(
      'w:val="single" w:sz="4" w:color="000000" w:frame="1"'
    );
    const xml = await emittedDocXml(buffer);
    expect(xml).toMatch(/<w:tblPrEx>[\s\S]*<w:top[^>]*w:frame="1"/);
  });
});
