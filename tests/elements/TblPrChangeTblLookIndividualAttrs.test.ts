/**
 * `<w:tblPrChange>` → previous `<w:tblLook>` — individual-attribute
 * format must round-trip.
 *
 * Per ECMA-376 §17.4.57 CT_TblLook accepts either a hex bitmask
 * (`w:val="04A0"`) or six individual CT_OnOff attributes (firstRow,
 * lastRow, firstColumn, lastColumn, noHBand, noVBand). Word commonly
 * emits the expanded individual-attribute form.
 *
 * `parseGenericPreviousProperties` (shared by `tblPrChange`,
 * `trPrChange`, and `tcPrChange` previous-property parsers) used a
 * hex-only read:
 *
 *   result.tblLook = look['@_w:val'] || '0000';
 *
 * That silently flattened every individually-set flag to `'0000'` on
 * round-trip. A track-changes scenario where the PREVIOUS tblLook
 * used the expanded form (very common in Word-authored documents)
 * lost the real flag set.
 *
 * Iteration 95 mirrors the main table-parser logic (same fix applied
 * to `parseTablePropertiesFromObject` in iteration 91) so both hex
 * and individual forms round-trip intact inside tblPrChange.
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

describe('<w:tblPrChange> previous <w:tblLook> round-trip', () => {
  it('preserves hex-form tblLook inside tblPrChange (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblLook w:val="04A0"/>
        <w:tblPrChange w:id="1" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:tblPr>
            <w:tblW w:w="4000" w:type="pct"/>
            <w:tblLook w:val="0620"/>
          </w:tblPr>
        </w:tblPrChange>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:tblPrChange[\s\S]*?<\/w:tblPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:tblLook[^>]*w:val="0620"/);
  });

  it('reconstructs hex tblLook from individual attrs inside tblPrChange', async () => {
    // Previous tblLook uses individual ST_OnOff attrs (no w:val).
    //   firstRow=1     → 0x0020 =   32
    //   firstColumn=1  → 0x0080 =  128
    //   noVBand=1      → 0x0400 = 1024
    // Sum = 1184 → hex "04A0"
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblLook w:val="0000"/>
        <w:tblPrChange w:id="2" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:tblPr>
            <w:tblW w:w="5000" w:type="pct"/>
            <w:tblLook w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
          </w:tblPr>
        </w:tblPrChange>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:tblPrChange[\s\S]*?<\/w:tblPrChange>/)?.[0] ?? '';
    // Previously this returned w:val="0000" because the hex-only reader
    // ignored the six individual attrs.
    expect(changeBlock).toMatch(/<w:tblLook[^>]*w:val="04A0"/);
  });

  it('honours ST_OnOff "off" literal in individual-attr tblLook', async () => {
    //   firstRow=on, noVBand=off → only 0x0020 = hex "0020"
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblLook w:val="0000"/>
        <w:tblPrChange w:id="3" w:author="Tester" w:date="2026-01-01T00:00:00Z">
          <w:tblPr>
            <w:tblW w:w="5000" w:type="pct"/>
            <w:tblLook w:firstRow="on" w:lastRow="off" w:firstColumn="off" w:lastColumn="off" w:noHBand="off" w:noVBand="off"/>
          </w:tblPr>
        </w:tblPrChange>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const changeBlock = out.match(/<w:tblPrChange[\s\S]*?<\/w:tblPrChange>/)?.[0] ?? '';
    expect(changeBlock).toMatch(/<w:tblLook[^>]*w:val="0020"/);
  });
});
