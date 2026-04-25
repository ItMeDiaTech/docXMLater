/**
 * Style-level rPr — character spacing, position, kerning, and language.
 *
 * The style-level rPr parser (`parseRunFormattingFromXml`) previously
 * skipped four commonly-used CT_RPr children even though the main run
 * parser and the `RunFormatting` interface both support them:
 *
 *   - `w:spacing` (§17.3.2.35, ST_SignedTwipsMeasure) → characterSpacing
 *   - `w:position` (§17.3.2.31, ST_SignedHpsMeasure) → position
 *   - `w:kern`     (§17.3.2.20, ST_HpsMeasure)        → kerning
 *   - `w:lang`     (§17.3.2.20, CT_Language)          → language
 *
 * A character style overriding any of these would be silently dropped on
 * parse, making programmatic resave lose the style override.
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
  <w:style w:type="character" w:styleId="RPrExtra">
    <w:name w:val="RPrExtra"/>
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
  return doc.getStylesManager().getStyle('RPrExtra')?.getRunFormatting();
}

describe('Style rPr — characterSpacing (w:spacing §17.3.2.35)', () => {
  it('parses <w:spacing w:val="20"/> as characterSpacing: 20', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:spacing w:val="20"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.characterSpacing).toBe(20);
    doc.dispose();
  });

  it('parses <w:spacing w:val="0"/> as characterSpacing: 0 (baseline reset)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:spacing w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.characterSpacing).toBe(0);
    doc.dispose();
  });

  it('parses <w:spacing w:val="-10"/> as characterSpacing: -10 (tighter)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:spacing w:val="-10"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.characterSpacing).toBe(-10);
    doc.dispose();
  });
});

describe('Style rPr — position (w:position §17.3.2.31)', () => {
  it('parses <w:position w:val="6"/> as position: 6 (raised)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:position w:val="6"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.position).toBe(6);
    doc.dispose();
  });

  it('parses <w:position w:val="-4"/> as position: -4 (lowered)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:position w:val="-4"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.position).toBe(-4);
    doc.dispose();
  });

  it('parses <w:position w:val="0"/> as position: 0 (explicit baseline)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:position w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.position).toBe(0);
    doc.dispose();
  });
});

describe('Style rPr — kerning (w:kern §17.3.2.20)', () => {
  it('parses <w:kern w:val="20"/> as kerning: 20', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:kern w:val="20"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.kerning).toBe(20);
    doc.dispose();
  });

  it('parses <w:kern w:val="0"/> as kerning: 0 (kern at every size)', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:kern w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.kerning).toBe(0);
    doc.dispose();
  });
});

describe('Style rPr — language (w:lang §17.3.2.20)', () => {
  it('parses <w:lang w:val="en-US"/> as language: "en-US"', async () => {
    const buffer = await makeDocxWithStyleRPr('<w:lang w:val="en-US"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getRPr(doc)?.language).toBe('en-US');
    doc.dispose();
  });

  it('parses multi-attribute <w:lang> as LanguageConfig object', async () => {
    const buffer = await makeDocxWithStyleRPr(
      '<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const lang = getRPr(doc)?.language;
    expect(typeof lang).toBe('object');
    if (typeof lang === 'object' && lang !== null) {
      expect(lang.val).toBe('en-US');
      expect(lang.eastAsia).toBe('zh-CN');
      expect(lang.bidi).toBe('ar-SA');
    }
    doc.dispose();
  });
});

describe('Style rPr — absent fields leave undefined', () => {
  it('leaves all four fields undefined when rPr is empty', async () => {
    const buffer = await makeDocxWithStyleRPr('');
    const doc = await Document.loadFromBuffer(buffer);
    const rPr = getRPr(doc);
    expect(rPr?.characterSpacing).toBeUndefined();
    expect(rPr?.position).toBeUndefined();
    expect(rPr?.kerning).toBeUndefined();
    expect(rPr?.language).toBeUndefined();
    doc.dispose();
  });
});
