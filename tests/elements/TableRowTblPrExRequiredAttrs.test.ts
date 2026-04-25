/**
 * Table row `<w:tblPrEx>` — required-attribute compliance for child elements.
 *
 * Per ECMA-376 Part 1, several children of `<w:tblPrEx>` (CT_TblPrEx, §17.4.62)
 * have REQUIRED attributes:
 *
 *   - `<w:shd>` (CT_Shd §17.3.1.31) — `w:val` (ST_Shd) is required; defines
 *     the shading pattern (clear, solid, pct5, pct10, …).
 *   - `<w:top>`/`<w:left>`/`<w:bottom>`/`<w:right>`/`<w:insideH>`/`<w:insideV>`
 *     (CT_Border §17.18.2) — `w:val` (ST_Border) is required; defines the
 *     border-line style (nil, single, thick, double, dotted, …).
 *
 * The inline `buildTablePropertyExceptionsXML` helper only emitted `w:val`
 * on `<w:shd>` when `shading.pattern` was truthy — so a user setting only
 * `fill: "FF0000"` produced `<w:shd w:fill="FF0000"/>` without `w:val`,
 * which fails strict OOXML schema validation. Same pattern on borders:
 * setting only `size`/`color` without `style` omitted `w:val`, failing
 * CT_Border validation.
 *
 * The shared `buildShadingAttributes` helper in `elements/CommonTypes.ts`
 * already defaults `w:val` to `"clear"` for this exact reason. The tblPrEx
 * emitter should use it (and default border style to `"nil"`) instead of
 * rolling its own partial copy.
 */

import { Table } from '../../src/elements/Table';
import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { XMLBuilder } from '../../src/xml/XMLBuilder';

describe('tblPrEx required-attribute compliance (§17.4.62 / §17.3.1.31 / §17.18.2)', () => {
  it('emits <w:shd w:val="clear" ...> when only fill is set (pattern missing)', () => {
    const table = new Table(2, 2);
    const row = table.getRows()[0]!;
    row.setTablePropertyExceptions({ shading: { fill: 'FFFFFF' } });
    const xml = XMLBuilder.elementToString(row.toXML());
    // The shd element must carry w:val. Default pattern is "clear".
    expect(xml).toMatch(/<w:shd[^>]*w:val="clear"[^>]*w:fill="FFFFFF"/);
    // Must not emit a shd element missing w:val entirely.
    expect(xml).not.toMatch(/<w:shd\s+w:fill="FFFFFF"\s*\/>/);
  });

  it('emits <w:top w:val="nil" ...> when border sets size/color without style', () => {
    const table = new Table(2, 2);
    const row = table.getRows()[0]!;
    row.setTablePropertyExceptions({
      borders: { top: { size: 4, color: '000000' } },
    });
    const xml = XMLBuilder.elementToString(row.toXML());
    // top border must carry w:val (default "nil" when style undefined).
    expect(xml).toMatch(
      /<w:tblBorders>[\s\S]*?<w:top[^>]*w:val="nil"[^>]*w:sz="4"[\s\S]*?<\/w:tblBorders>/
    );
    // Must not emit a bare <w:top w:sz="4" w:color="000000"/> without w:val.
    expect(xml).not.toMatch(/<w:top\s+w:sz="4"\s+w:color="000000"\s*\/>/);
  });

  it('passes OOXML validator after load/save round-trip with partial shd/border', async () => {
    // Build a doc directly in XML with tblPrEx children missing w:val.
    // Before the fix, round-trip through save validates against the SDK and
    // fails. After the fix, the generator supplies defaults and the doc
    // remains validator-clean.
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
          <w:shd w:val="clear" w:fill="FFFFFF"/>
          <w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/></w:tblBorders>
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
    const buffer = await zipHandler.toBuffer();
    const doc = await Document.loadFromBuffer(buffer);
    // Mutate so generator path re-emits tblPrEx from object model.
    const row = doc.getTables()[0]!.getRows()[0]!;
    const existing = row.getTablePropertyExceptions() ?? {};
    row.setTablePropertyExceptions({
      ...existing,
      // Strip `pattern`/`style` to simulate consumer code that only sets
      // color-ish fields — generator must still emit `w:val` defaults so
      // the validator accepts the result.
      shading: { fill: 'FFFFFF' },
      borders: { top: { size: 4, color: '000000' } },
    });
    // toBuffer runs the validator via the test setup. If the generator
    // omits required w:val, validation throws.
    await expect(doc.toBuffer()).resolves.toBeInstanceOf(Buffer);
    doc.dispose();
  });
});
