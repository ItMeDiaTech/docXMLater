/**
 * `<w:trHeight>` — zero-value row height must round-trip.
 *
 * Per ECMA-376 Part 1 §17.4.81 CT_Height:
 *   - `w:val` is ST_TwipsMeasure — zero is a valid value.
 *   - Combined with `w:hRule="exact"`, a zero-height row represents
 *     a hidden/collapsed row (used in some templates for layout-only
 *     spacers, or when cells are conditionally hidden).
 *
 * Bug: `parseTableRowPropertiesFromObject` (DocumentParser.ts §7123)
 * gated height storage with `if (heightVal > 0)`:
 *
 *     const heightVal = parseInt(trPrObj['w:trHeight']['@_w:val'] || '0', 10);
 *     if (heightVal > 0) { row.setHeight(heightVal); … }
 *
 * XMLParser coerces `"0"` to the number 0; the `> 0` gate drops it.
 * The emitter (`TableRow.ts §914`) already uses `!== undefined`, so
 * the parser/emitter asymmetry silently lost hidden-row height data
 * on every load → save cycle, and every tracked-change history that
 * recorded a previous zero height was collapsed to "no height set".
 *
 * Iteration 133 replaces the `> 0` filter with a zero-inclusive
 * `isExplicitlySet` gate on both the main-path parser.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function loadRowWith(trPrInnerXml: string) {
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
        <w:trPr>${trPrInnerXml}</w:trPr>
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
  return Document.loadFromBuffer(buffer);
}

describe('<w:trHeight w:val="0"/> zero-height row round-trip', () => {
  it('preserves explicit zero height with hRule="exact" (hidden row)', async () => {
    const doc = await loadRowWith('<w:trHeight w:val="0" w:hRule="exact"/>');
    const row = doc.getTables()[0]!.getRows()[0]!;
    const height = row.getHeight();
    const rule = row.getHeightRule();
    doc.dispose();
    expect(height).toBe(0);
    expect(rule).toBe('exact');
  });

  it('preserves explicit zero height with hRule="atLeast"', async () => {
    const doc = await loadRowWith('<w:trHeight w:val="0" w:hRule="atLeast"/>');
    const row = doc.getTables()[0]!.getRows()[0]!;
    const height = row.getHeight();
    doc.dispose();
    expect(height).toBe(0);
  });

  it('preserves zero height without hRule (defaults to auto)', async () => {
    const doc = await loadRowWith('<w:trHeight w:val="0"/>');
    const row = doc.getTables()[0]!.getRows()[0]!;
    const height = row.getHeight();
    doc.dispose();
    expect(height).toBe(0);
  });

  it('preserves positive height (regression guard)', async () => {
    const doc = await loadRowWith('<w:trHeight w:val="500" w:hRule="exact"/>');
    const row = doc.getTables()[0]!.getRows()[0]!;
    const height = row.getHeight();
    const rule = row.getHeightRule();
    doc.dispose();
    expect(height).toBe(500);
    expect(rule).toBe('exact');
  });

  it('omits height when trHeight element is absent (regression guard)', async () => {
    const doc = await loadRowWith('<w:tblHeader/>');
    const row = doc.getTables()[0]!.getRows()[0]!;
    const height = row.getHeight();
    doc.dispose();
    expect(height).toBeUndefined();
  });
});
