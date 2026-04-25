/**
 * Run-level numeric `@w:val` attributes — value-0 falsy-check bug.
 *
 * Three run-property parsers in DocumentParser used the pattern:
 *
 *     if (rPrObj['w:x']) {
 *       const val = rPrObj['w:x']['@_w:val'];
 *       if (val) run.setX(parseInt(val, 10));   // BUG: 0 is falsy
 *     }
 *
 * With XMLParser's `parseAttributeValue: true`, the attribute literal
 * "0" is coerced to the JS number `0`, which is falsy — so a Run with
 * `w:spacing w:val="0"` / `w:position w:val="0"` / `w:kern w:val="0"`
 * silently lost the formatting on load. Negative and positive values
 * worked (they're truthy), so the failure mode only shows up at the
 * boundary.
 *
 * Per ECMA-376 Part 1:
 *   §17.3.2.35 `w:spacing` — ST_SignedTwipsMeasure: 0 is the default /
 *     "no extra spacing", explicit.
 *   §17.3.2.31 `w:position` — ST_SignedHpsMeasure: 0 = baseline
 *     (no raise/lower, the default reset value).
 *   §17.3.2.20 `w:kern` — ST_HpsMeasure: 0 = "no kerning-size threshold"
 *     (kern glyphs at every size).
 *
 * All three are valid values that an explicit formatting setter or
 * a style override would reasonably emit. The rPrChange parser in
 * the same file (~5320) already used `!== undefined` correctly, so
 * the tracked-previous property round-tripped fine while the current
 * run dropped it — same asymmetry as iteration 19's tab-pos-0 bug.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithRunRPr(rPrXml: string): Promise<Buffer> {
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
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr>${rPrXml}</w:rPr><w:t>test</w:t></w:r></w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('Run parser — w:spacing (character spacing) honours value 0', () => {
  it('parses <w:spacing w:val="0"/> as 0', async () => {
    const buffer = await makeDocxWithRunRPr('<w:spacing w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.characterSpacing).toBe(0);
    doc.dispose();
  });

  it('parses negative <w:spacing w:val="-20"/> correctly (regression)', async () => {
    const buffer = await makeDocxWithRunRPr('<w:spacing w:val="-20"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.characterSpacing).toBe(-20);
    doc.dispose();
  });

  it('parses positive <w:spacing w:val="40"/> correctly (regression)', async () => {
    const buffer = await makeDocxWithRunRPr('<w:spacing w:val="40"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.characterSpacing).toBe(40);
    doc.dispose();
  });
});

describe('Run parser — w:position (vertical position) honours value 0', () => {
  it('parses <w:position w:val="0"/> as 0 (baseline)', async () => {
    const buffer = await makeDocxWithRunRPr('<w:position w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.position).toBe(0);
    doc.dispose();
  });

  it('parses negative <w:position w:val="-6"/> correctly (regression)', async () => {
    const buffer = await makeDocxWithRunRPr('<w:position w:val="-6"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.position).toBe(-6);
    doc.dispose();
  });

  it('parses positive <w:position w:val="4"/> correctly (regression)', async () => {
    const buffer = await makeDocxWithRunRPr('<w:position w:val="4"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.position).toBe(4);
    doc.dispose();
  });
});

describe('Run parser — w:kern (kerning threshold) honours value 0', () => {
  it('parses <w:kern w:val="0"/> as 0 (kern at all sizes)', async () => {
    const buffer = await makeDocxWithRunRPr('<w:kern w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.kerning).toBe(0);
    doc.dispose();
  });

  it('parses positive <w:kern w:val="24"/> correctly (regression)', async () => {
    const buffer = await makeDocxWithRunRPr('<w:kern w:val="24"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const fmt = doc.getParagraphs()[0]!.getRuns()[0]!.getFormatting();
    expect(fmt.kerning).toBe(24);
    doc.dispose();
  });
});

describe('Run parser — full round-trip of value-0 numeric attributes', () => {
  it('round-trips w:spacing="0" through load → save → load', async () => {
    const buffer1 = await makeDocxWithRunRPr('<w:spacing w:val="0"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(doc1.getParagraphs()[0]!.getRuns()[0]!.getFormatting().characterSpacing).toBe(0);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(doc2.getParagraphs()[0]!.getRuns()[0]!.getFormatting().characterSpacing).toBe(0);
    doc2.dispose();
  });

  it('round-trips w:position="0" through load → save → load', async () => {
    const buffer1 = await makeDocxWithRunRPr('<w:position w:val="0"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(doc1.getParagraphs()[0]!.getRuns()[0]!.getFormatting().position).toBe(0);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(doc2.getParagraphs()[0]!.getRuns()[0]!.getFormatting().position).toBe(0);
    doc2.dispose();
  });
});
