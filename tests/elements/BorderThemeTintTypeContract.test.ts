/**
 * Border `w:themeTint` / `w:themeShade` — numeric-looking hex values
 * must round-trip as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1 §17.18.82 CT_Border, `w:themeTint` and
 * `w:themeShade` are `ST_UcharHexNumber` (2-character hex). The same
 * applies to CT_TopBorder / CT_BottomBorder / CT_Border extensions
 * used on table, cell, run, and page borders. All these are declared
 * as `string` on the TypeScript interfaces.
 *
 * Bug: XMLParser's `parseAttributeValue: true` coerces purely-digit
 * hex strings like `"80"` / `"50"` / `"40"` to JS numbers. Four
 * object-form border parsers stored the raw coerced value:
 *   - Run border `<w:bdr>` (`DocumentParser.ts:5185`)
 *   - rPrChange previous `<w:bdr>` (`DocumentParser.ts:5731`)
 *   - Shared table/cell `parseBorderElement` (used by main tbl, main
 *     tc, tblPrEx, tblPrChange, tcPrChange — `DocumentParser.ts:6571`)
 *   - `parseTableBordersFromObject` (used by tblPrEx + the generic
 *     *PrChange previous-properties parser — `DocumentParser.ts:7268`)
 *
 * Like the shading fix in iter 123, the numeric leak only surfaced
 * when downstream code called string methods on the stored value
 * (`.toUpperCase()`, `.startsWith(...)` etc.) — on the wire the
 * emitter's `XMLBuilder.wSelf` coerces numbers to strings so round-
 * trip XML looked correct.
 *
 * Iteration 124 casts every themed border attribute through
 * `String(...)` in all four object-form parsers.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { Table } from '../../src/elements/Table';

async function loadTableWithBorderAttr(
  tblBorderInner: string
): Promise<{ themeTint?: unknown; themeShade?: unknown }> {
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
        <w:tblBorders>
          ${tblBorderInner}
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
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  const doc = await Document.loadFromBuffer(buffer);
  const table = doc.getTables()[0] as Table;
  const top = table.getFormatting().borders?.top as {
    themeTint?: unknown;
    themeShade?: unknown;
  };
  doc.dispose();
  return top ?? {};
}

describe('<w:tblBorders> themeTint/themeShade type-contract preservation', () => {
  it('stores w:themeTint="80" as the STRING "80" (not number 80)', async () => {
    const top = await loadTableWithBorderAttr(
      '<w:top w:val="single" w:sz="4" w:color="auto" w:themeColor="accent1" w:themeTint="80"/>'
    );
    expect(typeof top.themeTint).toBe('string');
    expect(top.themeTint).toBe('80');
  });

  it('stores w:themeShade="50" as the STRING "50"', async () => {
    const top = await loadTableWithBorderAttr(
      '<w:top w:val="single" w:sz="4" w:color="auto" w:themeColor="accent1" w:themeShade="50"/>'
    );
    expect(typeof top.themeShade).toBe('string');
    expect(top.themeShade).toBe('50');
  });

  it('preserves non-numeric hex like "FF" as string (regression guard)', async () => {
    const top = await loadTableWithBorderAttr(
      '<w:top w:val="single" w:sz="4" w:color="auto" w:themeColor="accent1" w:themeTint="FF"/>'
    );
    expect(typeof top.themeTint).toBe('string');
    expect(top.themeTint).toBe('FF');
  });

  it('string methods are callable on parsed themeTint', async () => {
    const top = await loadTableWithBorderAttr(
      '<w:top w:val="single" w:sz="4" w:color="auto" w:themeColor="accent1" w:themeTint="80"/>'
    );
    const tint = top.themeTint as string | undefined;
    expect(() => tint?.toUpperCase()).not.toThrow();
    expect(tint?.toUpperCase()).toBe('80');
  });
});
