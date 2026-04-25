/**
 * `parseGenericPreviousProperties` — ST_String attributes inside
 * `w:tblPrChange` / `w:trPrChange` / `w:tcPrChange` previousProperties
 * must round-trip as strings (type-contract safety).
 *
 * Per ECMA-376 Part 1:
 *   - `w:tblStyle` w:val is ST_String (§17.7.4.62) — style reference
 *   - `w:tblCaption` w:val is ST_String (§17.4.62) — accessible caption
 *   - `w:tblDescription` w:val is ST_String (§17.4.63) — accessible descr
 *   - `w:cnfStyle` w:val is ST_Cnf (§17.4.7) — 12-character binary
 *     bitmask string like `"100000000000"` (firstRow=first bit, etc.)
 *
 * Bug: XMLParser's `parseAttributeValue: true` coerces purely-digit
 * strings to JS numbers. Custom style IDs (`"2025"`), numeric captions
 * (`"42"`), and — most critically — the `w:cnfStyle` bitmask (always a
 * run of `0` and `1` digits, e.g. `"100000000000"`) are all coerced.
 *
 * `parseGenericPreviousProperties` previously stored the raw coerced
 * value, so the tracked-change history would report a `number` where
 * the interface declares a `string`. Any downstream code that called
 * `.charAt(i)` / `.startsWith(...)` on the previous style or cnfStyle
 * bitmask would throw.
 *
 * Iteration 128 casts all four ST_String previous-property reads
 * through `String(...)`.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadDocWith(tableBodyXml: string) {
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
    ${tableBodyXml}
    <w:p><w:r><w:t>trailer</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  const buffer = await zipHandler.toBuffer();
  return Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
}

const tableWithTblPrChange = (tblStyleVal: string, caption: string, description: string) => `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="5000" w:type="pct"/>
        <w:tblPrChange w:id="1" w:author="A" w:date="2026-04-24T10:00:00Z">
          <w:tblPr>
            <w:tblStyle w:val="${tblStyleVal}"/>
            <w:tblW w:w="3000" w:type="pct"/>
            <w:tblCaption w:val="${caption}"/>
            <w:tblDescription w:val="${description}"/>
          </w:tblPr>
        </w:tblPrChange>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>x</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const tableWithTrPrChangeCnfStyle = (cnfStyle: string) => `
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr>
          <w:cnfStyle w:val="100000000000"/>
          <w:trPrChange w:id="2" w:author="A" w:date="2026-04-24T10:00:00Z">
            <w:trPr>
              <w:cnfStyle w:val="${cnfStyle}"/>
              <w:trHeight w:val="500"/>
            </w:trPr>
          </w:trPrChange>
        </w:trPr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>
          <w:p><w:r><w:t>row</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

describe('*PrChange previousProperties ST_String type-contract preservation', () => {
  it('stores tblPrChange <w:tblStyle w:val="2025"/> as STRING "2025"', async () => {
    const doc = await loadDocWith(tableWithTblPrChange('2025', 'cap', 'desc'));
    const table = doc.getTables()[0]!;
    const tblPrChange = table.getTblPrChange();
    doc.dispose();
    expect(tblPrChange).toBeDefined();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prev = tblPrChange!.previousProperties as any;
    expect(typeof prev.style).toBe('string');
    expect(prev.style).toBe('2025');
  });

  it('stores tblPrChange <w:tblCaption w:val="42"/> as STRING "42"', async () => {
    const doc = await loadDocWith(tableWithTblPrChange('TableGrid', '42', 'desc'));
    const table = doc.getTables()[0]!;
    const tblPrChange = table.getTblPrChange();
    doc.dispose();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prev = tblPrChange!.previousProperties as any;
    expect(typeof prev.caption).toBe('string');
    expect(prev.caption).toBe('42');
  });

  it('stores tblPrChange <w:tblDescription w:val="99"/> as STRING "99"', async () => {
    const doc = await loadDocWith(tableWithTblPrChange('TableGrid', 'cap', '99'));
    const table = doc.getTables()[0]!;
    const tblPrChange = table.getTblPrChange();
    doc.dispose();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prev = tblPrChange!.previousProperties as any;
    expect(typeof prev.description).toBe('string');
    expect(prev.description).toBe('99');
  });

  it('stores trPrChange <w:cnfStyle w:val="100000000000"/> as STRING "100000000000"', async () => {
    const doc = await loadDocWith(tableWithTrPrChangeCnfStyle('100000000000'));
    const table = doc.getTables()[0]!;
    const row = table.getRows()[0]!;
    const trPrChange = row.getTrPrChange();
    doc.dispose();
    expect(trPrChange).toBeDefined();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prev = trPrChange!.previousProperties as any;
    expect(typeof prev.cnfStyle).toBe('string');
    expect(prev.cnfStyle).toBe('100000000000');
  });

  it('cnfStyle bitmask supports .charAt() after parse (no number coercion)', async () => {
    // cnfStyle layout per ECMA-376 §17.4.7: first bit = firstRow
    // The string "100000000000" means "firstRow=on, all others=off"
    const doc = await loadDocWith(tableWithTrPrChangeCnfStyle('010000000010'));
    const table = doc.getTables()[0]!;
    const row = table.getRows()[0]!;
    const trPrChange = row.getTrPrChange();
    doc.dispose();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const cnf = (trPrChange!.previousProperties as any).cnfStyle as string;
    // Pre-fix: cnfStyle was number 10000000010, .charAt would throw.
    expect(() => cnf.charAt(0)).not.toThrow();
    expect(cnf.length).toBe(12);
    expect(cnf.charAt(1)).toBe('1');
    expect(cnf.charAt(10)).toBe('1');
  });

  it('preserves non-numeric tblStyle "Heading1" (regression guard)', async () => {
    const doc = await loadDocWith(tableWithTblPrChange('Heading1', 'cap', 'desc'));
    const table = doc.getTables()[0]!;
    const tblPrChange = table.getTblPrChange();
    doc.dispose();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const prev = tblPrChange!.previousProperties as any;
    expect(prev.style).toBe('Heading1');
  });
});
