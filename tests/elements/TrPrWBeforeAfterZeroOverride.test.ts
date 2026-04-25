/**
 * Row-level `<w:wBefore>`, `<w:wAfter>`, and `<w:tblCellSpacing>` —
 * zero-value overrides must round-trip.
 *
 * Per ECMA-376 Part 1:
 *   - §17.4.83 w:wBefore, §17.4.82 w:wAfter — both wrap `w:w` / `w:type`
 *     (ST_TblWidth). `w:w="0"` paired with `w:type="auto"` is the
 *     idiomatic "empty gutter" form; `w:w="0"` in `dxa` twips can
 *     explicitly override an inherited non-zero width.
 *   - §17.4.44 w:tblCellSpacing (row-level) — same ST_TblWidth
 *     semantics. `w:w="0"` is an explicit "no cell spacing" override
 *     of a table-level tblCellSpacing.
 *
 * The parser (`parseTableRowPropertiesFromObject`) guarded each with
 * `if (w > 0)`, silently dropping zero-value overrides on load. The
 * emitter uses `!== undefined`, so the parser/emitter asymmetry meant
 * any row relying on an explicit zero override silently reverted to
 * inheriting the table-level value on every round-trip.
 *
 * Iteration 105 swaps the `> 0` gates for `isExplicitlySet` +
 * `safeParseInt` across wBefore, wAfter, and tblCellSpacing — same
 * approach as the iter 104 fix for `<w:tblPrEx>`.
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

function buildRowDoc(trPrInner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblCellSpacing w:w="100" w:type="dxa"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr>${trPrInner}</w:trPr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`;
}

describe('<w:wBefore> / <w:wAfter> / <w:tblCellSpacing> zero-override round-trip', () => {
  it('preserves w:wBefore w:w="0" (explicit zero gutter before row)', async () => {
    const out = await loadAndResaveDocXml(buildRowDoc('<w:wBefore w:w="0" w:type="auto"/>'));
    expect(out).toMatch(/<w:wBefore[^/]*w:w="0"/);
  });

  it('preserves w:wAfter w:w="0" (explicit zero gutter after row)', async () => {
    const out = await loadAndResaveDocXml(buildRowDoc('<w:wAfter w:w="0" w:type="auto"/>'));
    expect(out).toMatch(/<w:wAfter[^/]*w:w="0"/);
  });

  it('preserves w:tblCellSpacing w:w="0" (row-level override of table spacing)', async () => {
    const out = await loadAndResaveDocXml(buildRowDoc('<w:tblCellSpacing w:w="0" w:type="dxa"/>'));
    // Extract the ROW-level tblCellSpacing (nested in trPr), not the
    // table-level one in tblPr above it.
    const rowTblPr = out.match(/<w:trPr>[\s\S]*?<\/w:trPr>/)?.[0] ?? '';
    expect(rowTblPr).toMatch(/<w:tblCellSpacing[^/]*w:w="0"/);
  });

  it('still preserves positive values (regression guard)', async () => {
    const out = await loadAndResaveDocXml(
      buildRowDoc('<w:wBefore w:w="720" w:type="dxa"/><w:wAfter w:w="720" w:type="dxa"/>')
    );
    expect(out).toMatch(/<w:wBefore[^/]*w:w="720"/);
    expect(out).toMatch(/<w:wAfter[^/]*w:w="720"/);
  });
});
