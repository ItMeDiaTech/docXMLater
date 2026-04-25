/**
 * Main-path `<w:tblBorders>` / `<w:tcBorders>` — bidi-aware
 * `<w:start>` / `<w:end>` border aliases must parse on table-direct
 * and cell-direct borders (not just on style-level borders).
 *
 * Per ECMA-376 Part 1 §17.4.40 CT_TblBorders and §17.4.66 CT_TcBorders,
 * the left and right borders each have bidi-aware aliases:
 *   - `<w:start>` ≡ preferred bidi spelling of `<w:left>`
 *   - `<w:end>` ≡ preferred bidi spelling of `<w:right>`
 *
 * Modern Word documents (2013+) and Google Docs emit the bidi-aware
 * form by default. Iteration 121 fixed the style-level parser
 * (`parseBordersFromXml`). This iteration covers the three remaining
 * object-form parsers:
 *   - Main table borders in `parseTableFromObject` (line ~6926)
 *   - Main cell borders in `parseTableCellFromObject` (line ~7338)
 *   - Shared `parseTableBordersFromObject` + `parseGenericPreviousProperties`
 *     tblBorders/tcBorders (used by tblPrEx, tblPrChange, tcPrChange)
 *
 * Before this fix, any direct table/cell authored with `<w:start>` /
 * `<w:end>` silently lost the side borders on load. The emitter writes
 * legacy `w:left` / `w:right`, so the bidi-aware input had no counter-
 * part to flow to on re-save — authored side borders disappeared.
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

describe('table/cell borders bidi-aware <w:start>/<w:end> on main parsers', () => {
  it('main table tblBorders: <w:start>/<w:end> round-trip as w:left/w:right', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:start w:val="double" w:sz="8" w:color="FF0000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:end w:val="double" w:sz="8" w:color="00FF00"/>
        </w:tblBorders>
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
    // The emitter always writes w:left/w:right (legacy keys). The
    // parser should have loaded the values so those appear in output.
    const tblBorders = out.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/)?.[0] ?? '';
    expect(tblBorders).toMatch(/<w:left[^/]*w:val="double"/);
    expect(tblBorders).toMatch(/<w:left[^/]*w:color="FF0000"/);
    expect(tblBorders).toMatch(/<w:right[^/]*w:val="double"/);
    expect(tblBorders).toMatch(/<w:right[^/]*w:color="00FF00"/);
  });

  it('main cell tcBorders: <w:start>/<w:end> round-trip as w:left/w:right', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            <w:tcBorders>
              <w:top w:val="single" w:sz="4" w:color="000000"/>
              <w:start w:val="dashed" w:sz="6" w:color="0000FF"/>
              <w:bottom w:val="single" w:sz="4" w:color="000000"/>
              <w:end w:val="dashed" w:sz="6" w:color="FF00FF"/>
            </w:tcBorders>
          </w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    const out = await loadAndResaveDocXml(xml);
    const tcBorders = out.match(/<w:tcBorders>[\s\S]*?<\/w:tcBorders>/)?.[0] ?? '';
    expect(tcBorders).toMatch(/<w:left[^/]*w:val="dashed"/);
    expect(tcBorders).toMatch(/<w:left[^/]*w:color="0000FF"/);
    expect(tcBorders).toMatch(/<w:right[^/]*w:val="dashed"/);
    expect(tcBorders).toMatch(/<w:right[^/]*w:color="FF00FF"/);
  });

  it('still parses legacy w:left/w:right (regression guard)', async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="000000"/>
          <w:left w:val="dotted" w:sz="2" w:color="AA0000"/>
          <w:bottom w:val="single" w:sz="4" w:color="000000"/>
          <w:right w:val="dotted" w:sz="2" w:color="00AA00"/>
        </w:tblBorders>
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
    const tblBorders = out.match(/<w:tblBorders>[\s\S]*?<\/w:tblBorders>/)?.[0] ?? '';
    expect(tblBorders).toMatch(/<w:left[^/]*w:val="dotted"/);
    expect(tblBorders).toMatch(/<w:right[^/]*w:val="dotted"/);
  });
});
