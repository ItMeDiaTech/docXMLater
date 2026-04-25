/**
 * `<w:shd>` — numeric-looking attribute values like `w:themeFillTint="80"`
 * must round-trip as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1 §17.3.1.32 CT_Shd, every string-typed attribute
 * on a shading element is declared as either `ST_UcharHexNumber` (2-char
 * hex tint / shade), `ST_ThemeColor` (enum), `ST_HexColor` (6-char hex
 * + "auto"), or `ST_Shd` (pattern enum). All of these map to
 * `string` on the `ShadingConfig` interface.
 *
 * Bug: XMLParser's `parseAttributeValue: true` coerces purely-digit
 * hex strings like `"80"` to the JS number `80`. The
 * `parseShadingFromObj` parser stored the raw coerced value
 * (`shading.themeFillTint = 80` — a number in a `string` field). On
 * emit, the `buildShadingAttributes` helper just forwarded the value
 * to `XMLBuilder.wSelf`, which coerces numbers to strings for XML
 * output — so the round-trip looked fine on the wire. But downstream
 * code that expected a string (e.g., `.toUpperCase()`, `.startsWith(...)`
 * for hex-case normalisation) would throw or silently misbehave.
 *
 * Iteration 123 casts every shading string attribute through
 * `String(...)` in `parseShadingFromObj` so the declared TypeScript
 * contract holds.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadCellShading(shdXml: string): Promise<{ themeFillTint?: unknown }> {
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
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="5000" w:type="pct"/>
            ${shdXml}
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
  const table = doc.getTables()[0]!;
  const cell = table.getRows()[0]!.getCells()[0]!;
  const shading = cell.getShading();
  doc.dispose();
  return shading ?? {};
}

describe('<w:shd> numeric attribute type-contract preservation', () => {
  it('stores w:themeFillTint="80" as the STRING "80" (not number 80)', async () => {
    const shading = await loadCellShading(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFillTint="80"/>'
    );
    expect(typeof shading.themeFillTint).toBe('string');
    expect(shading.themeFillTint).toBe('80');
  });

  it('stores w:themeFillShade="50" as the STRING "50"', async () => {
    const shading = (await loadCellShading(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFillShade="50"/>'
    )) as { themeFillShade?: unknown };
    expect(typeof shading.themeFillShade).toBe('string');
    expect(shading.themeFillShade).toBe('50');
  });

  it('preserves w:themeFillTint="FF" (non-numeric hex) as string', async () => {
    const shading = await loadCellShading(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFillTint="FF"/>'
    );
    expect(typeof shading.themeFillTint).toBe('string');
    expect(shading.themeFillTint).toBe('FF');
  });

  it('round-trips string methods are callable on parsed tint', async () => {
    const shading = await loadCellShading(
      '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFillTint="80"/>'
    );
    // The real failure mode of the pre-fix bug: calling .toUpperCase()
    // on a number would throw. With the fix, the value IS a string, so
    // string methods work.
    const tint = shading.themeFillTint as string | undefined;
    expect(() => tint?.toUpperCase()).not.toThrow();
    expect(tint?.toUpperCase()).toBe('80');
  });
});
