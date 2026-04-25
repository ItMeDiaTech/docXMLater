/**
 * `<w:tc><w:tcPr><w:tcW w:w="0" w:type="…"/>` — table cell explicit
 * zero-width override must round-trip.
 *
 * Per ECMA-376 Part 1 §17.4.72 CT_TblWidth, `w:w` is
 * ST_MeasurementOrPercent and `w:type` is ST_TblWidth. A `w:w="0"`
 * value is legal in every type context:
 *   - `w:type="auto"` — the idiomatic "size to content" form (also
 *     the default when `w:tcW` is absent).
 *   - `w:type="dxa"` / `"pct"` / `"nil"` — an explicit "reset to
 *     zero" override of an inherited width from the table's cell
 *     defaults (`<w:tblCellMar>` / tblPr `w:tcW`) or a conditional
 *     cell style.
 *
 * The main cell parser at `DocumentParser.ts:7273` used
 *
 *     if (widthVal > 0 || widthType === 'auto') { cell.setWidthType(...) }
 *
 * so a cell with `w:w="0" w:type="dxa"` (or "pct"/"nil") was silently
 * dropped on load — the cell reverted to inheriting the ancestor
 * width. The emitter at `TableCell.ts:1353` uses `!== undefined`, so
 * the parser/emitter asymmetry broke round-trip on every zero-value
 * override.
 *
 * Iteration 118 swaps the `> 0` gate for `isExplicitlySet` +
 * `safeParseInt`, preserving every zero-value cell-width override.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Table } from '../../src/elements/Table';

async function loadAndReadFirstCellWidth(
  tcWAttrs: string
): Promise<{ width?: number; widthType?: string }> {
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
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr>
            <w:tcW ${tcWAttrs}/>
          </w:tcPr>
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
  const cell = table.getRows()[0]!.getCells()[0]!;
  const result = { width: cell.getWidth(), widthType: cell.getWidthType() };
  doc.dispose();
  return result;
}

describe('<w:tcW w:w="0" …> explicit zero-width cell override round-trip', () => {
  it('preserves w:w="0" w:type="dxa" (explicit "reset to zero twips")', async () => {
    const { width, widthType } = await loadAndReadFirstCellWidth('w:w="0" w:type="dxa"');
    expect(width).toBe(0);
    expect(widthType).toBe('dxa');
  });

  it('preserves w:w="0" w:type="pct" (explicit "0% width" override)', async () => {
    const { width, widthType } = await loadAndReadFirstCellWidth('w:w="0" w:type="pct"');
    expect(width).toBe(0);
    expect(widthType).toBe('pct');
  });

  it('preserves w:w="0" w:type="auto" (regression guard — already worked)', async () => {
    const { width, widthType } = await loadAndReadFirstCellWidth('w:w="0" w:type="auto"');
    expect(width).toBe(0);
    expect(widthType).toBe('auto');
  });

  it('preserves positive w:w values (regression guard)', async () => {
    const { width, widthType } = await loadAndReadFirstCellWidth('w:w="5000" w:type="pct"');
    expect(width).toBe(5000);
    expect(widthType).toBe('pct');
  });
});
