/**
 * Style-level rPr — scaling (w:w), emphasis (w:em), and effect (w:effect).
 *
 * Three commonly-used CT_RPr children that the style-level rPr parser
 * (`parseRunFormattingFromXml`) previously dropped:
 *
 *   - `w:w`      (§17.3.2.43, ST_TextScale)    → scaling  (percentage)
 *   - `w:em`     (§17.3.2.13, ST_Em)           → emphasis (dot / circle / ...)
 *   - `w:effect` (§17.3.2.12, ST_TextEffect)   → effect   (blink / shimmer / ...)
 *
 * Main run parser handles all three (DocumentParser.ts ~5072-5136) but
 * the style-level parser never touched them — character styles setting
 * any of these were silently stripped on load.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyleRPr(rPrInner: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();
  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );
  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="character" w:styleId="ScaleEmEff">
    <w:name w:val="ScaleEmEff"/>
    <w:rPr>${rPrInner}</w:rPr>
  </w:style>
</w:styles>`
  );
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

function getRPr(doc: Document) {
  return doc.getStylesManager().getStyle('ScaleEmEff')?.getRunFormatting();
}

describe('Style rPr — scaling (w:w §17.3.2.43)', () => {
  it('parses <w:w w:val="200"/> as scaling: 200 (200% width)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:w w:val="200"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.scaling).toBe(200);
    doc.dispose();
  });

  it('parses <w:w w:val="50"/> as scaling: 50 (half width)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:w w:val="50"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.scaling).toBe(50);
    doc.dispose();
  });
});

describe('Style rPr — emphasis (w:em §17.3.2.13)', () => {
  const MARKS = ['dot', 'comma', 'circle', 'underDot'] as const;
  for (const mark of MARKS) {
    it(`parses <w:em w:val="${mark}"/> as emphasis: "${mark}"`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:em w:val="${mark}"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      expect(getRPr(doc)?.emphasis).toBe(mark);
      doc.dispose();
    });
  }
});

describe('Style rPr — effect (w:effect §17.3.2.12)', () => {
  const EFFECTS = [
    'blinkBackground',
    'lights',
    'antsBlack',
    'antsRed',
    'shimmer',
    'sparkle',
  ] as const;
  for (const effect of EFFECTS) {
    it(`parses <w:effect w:val="${effect}"/> as effect: "${effect}"`, async () => {
      const buffer = await makeDocxWithStyleRPr(`<w:effect w:val="${effect}"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      expect(getRPr(doc)?.effect).toBe(effect);
      doc.dispose();
    });
  }
});

describe('Style rPr — absent fields leave undefined', () => {
  it('leaves scaling/emphasis/effect undefined when rPr is empty', async () => {
    const buffer = await makeDocxWithStyleRPr('');
    const doc = await Document.loadFromBuffer(buffer);
    const rPr = getRPr(doc);
    expect(rPr?.scaling).toBeUndefined();
    expect(rPr?.emphasis).toBeUndefined();
    expect(rPr?.effect).toBeUndefined();
    doc.dispose();
  });
});
