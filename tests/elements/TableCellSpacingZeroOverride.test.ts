/**
 * Table-level `<w:tblPr><w:tblCellSpacing w:w="0" …/>` — explicit
 * zero-spacing override must round-trip.
 *
 * Per ECMA-376 Part 1 §17.4.44 CT_TblCellSpacing, `w:w` is
 * ST_MeasurementOrPercent. `w:w="0"` is a legal "explicit zero cell
 * spacing" value — commonly used to override a style-inherited
 * non-zero tblCellSpacing back to no extra spacing.
 *
 * The main table parser (`DocumentParser.ts:6672`) guarded with
 * `if (spacing > 0)`, dropping the zero override on load. The emitter
 * at `Table.ts:1711` uses `!== undefined`, so the asymmetry meant a
 * table that explicitly reset tblCellSpacing to 0 silently reverted
 * to inheriting the style-level non-zero value on every round-trip.
 * This also caused tracked table-property changes (tblPrChange) to
 * capture wrong "previous" cellSpacing values when the previous state
 * used the zero form.
 *
 * Iteration 107 swaps the `> 0` gate for
 * `isExplicitlySet` + `safeParseInt`.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Table } from '../../src/elements/Table';

async function loadAndReadTableCellSpacing(
  valAttr: string | null
): Promise<{ value?: number; type?: string }> {
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
  const tblCellSpacing =
    valAttr === null ? '' : `<w:tblCellSpacing w:w="${valAttr}" w:type="dxa"/>`;
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        ${tblCellSpacing}
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
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const table = doc.getTables()[0] as Table;
  const result = { value: table.getCellSpacing(), type: table.getCellSpacingType() };
  doc.dispose();
  return result;
}

describe('<w:tblCellSpacing w:w="0"/> zero-value override round-trip', () => {
  it('preserves w:w="0" as an explicit zero override (previously dropped)', async () => {
    const { value, type } = await loadAndReadTableCellSpacing('0');
    expect(value).toBe(0);
    expect(type).toBe('dxa');
  });

  it('preserves positive values (regression guard)', async () => {
    const { value, type } = await loadAndReadTableCellSpacing('200');
    expect(value).toBe(200);
    expect(type).toBe('dxa');
  });

  it('returns undefined when tblCellSpacing element is absent (regression guard)', async () => {
    const { value } = await loadAndReadTableCellSpacing(null);
    expect(value).toBeUndefined();
  });
});
