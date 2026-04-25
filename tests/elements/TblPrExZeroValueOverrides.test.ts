/**
 * `<w:tblPrEx>` row-level exception values — zero-value (and negative)
 * overrides must round-trip.
 *
 * Per ECMA-376 Part 1 §17.4.61 CT_TblPrEx, w:tblW / w:tblCellSpacing /
 * w:tblInd each carry a ST_MeasurementOrPercent `w:w` attribute. Zero
 * and negative values are legal:
 *   - w:tblW="0" (paired with type="nil"/"auto") — explicit no-width
 *     override on a row that would otherwise inherit a width from the
 *     table-level tblW.
 *   - w:tblCellSpacing="0" — explicit "no extra spacing" override.
 *   - w:tblInd="-720" — outdent into the page margin.
 *
 * The parser previously guarded each with `if (val > 0)`, silently
 * dropping zero AND negative overrides on load. The emitter uses
 * `!== undefined`, so any round-tripped document lost the explicit
 * overrides — the row inherited the table-level value again.
 *
 * Iteration 104 swaps the `> 0` gates for `isExplicitlySet` +
 * `safeParseInt`, matching the emitter's sign-preserving semantics.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadAndReadFirstRowTblPrEx(rowTblPrExInner: string) {
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
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblCellSpacing w:w="100" w:type="dxa"/>
        <w:tblInd w:w="200" w:type="dxa"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tblPrEx>
          ${rowTblPrExInner}
        </w:tblPrEx>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>c</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:t>d</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const table = doc.getTables()[0]!;
  const row = table.getRows()[0]!;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const exceptions = (row as any).getTablePropertyExceptions?.();
  doc.dispose();
  return exceptions;
}

describe('<w:tblPrEx> zero / negative override round-trip', () => {
  it('preserves w:tblInd w:w="0" (explicit "no row indent" override)', async () => {
    const excs = await loadAndReadFirstRowTblPrEx('<w:tblInd w:w="0" w:type="dxa"/>');
    expect(excs?.indentation).toBe(0);
  });

  it('preserves w:tblInd w:w="-720" (outdent into margin)', async () => {
    const excs = await loadAndReadFirstRowTblPrEx('<w:tblInd w:w="-720" w:type="dxa"/>');
    expect(excs?.indentation).toBe(-720);
  });

  it('preserves w:tblCellSpacing w:w="0" (explicit "no cell spacing")', async () => {
    const excs = await loadAndReadFirstRowTblPrEx('<w:tblCellSpacing w:w="0" w:type="dxa"/>');
    expect(excs?.cellSpacing).toBe(0);
  });

  it('preserves w:tblW w:w="0" (explicit zero-width override)', async () => {
    const excs = await loadAndReadFirstRowTblPrEx('<w:tblW w:w="0" w:type="auto"/>');
    expect(excs?.width).toBe(0);
  });

  it('still parses positive overrides (regression guard)', async () => {
    const excs = await loadAndReadFirstRowTblPrEx(
      '<w:tblW w:w="3000" w:type="dxa"/><w:tblInd w:w="500" w:type="dxa"/>'
    );
    expect(excs?.width).toBe(3000);
    expect(excs?.indentation).toBe(500);
  });

  it('returns undefined when tblPrEx has no relevant children', async () => {
    const excs = await loadAndReadFirstRowTblPrEx('');
    // No exception data → the row should report undefined.
    // (Zero-length exceptions object → tablePropertyExceptions stays undefined.)
    expect(excs).toBeUndefined();
  });
});
