/**
 * Style-level `<w:spacing>` — full CT_Spacing attribute round-trip.
 *
 * Per ECMA-376 Part 1 §17.3.1.33, `<w:spacing>` inside a style's
 * `<w:pPr>` carries EIGHT attributes, not four:
 *
 *   w:before             — twips before paragraph
 *   w:beforeLines        — hundredths of a line, before (CJK-friendly)
 *   w:beforeAutospacing  — ST_OnOff: auto-calc before (overrides w:before)
 *   w:after              — twips after paragraph
 *   w:afterLines         — hundredths of a line, after
 *   w:afterAutospacing   — ST_OnOff: auto-calc after
 *   w:line               — line spacing value
 *   w:lineRule           — ST_LineSpacingRule: auto | exact | atLeast
 *
 * Bug this suite guards against:
 *   - `Style.toXML()` emitted only w:before / w:after / w:line / w:lineRule.
 *     The four line-unit / auto-spacing attributes were silently dropped
 *     on save even though `DocumentParser.parseParagraphFormattingFromXml`
 *     already correctly read all eight. So a style authored in a CJK
 *     locale with line-unit spacing (`w:beforeLines="100"`) or with
 *     Word's auto-spacing flags (`w:beforeAutospacing="1"`) would
 *     round-trip with those attributes dropped, silently switching
 *     the style's spacing behaviour on every save.
 *
 *   Matches the pattern from iteration 23's `<w:ind>` work — the
 *   style-level code path was the one missing coverage after the
 *   main-path and pPrChange paths got fixed earlier.
 */

import { Document } from '../../src/core/Document';
import { Style } from '../../src/formatting/Style';
import { XMLBuilder } from '../../src/xml/XMLBuilder';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStylePPr(pPrInner: string): Promise<Buffer> {
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
  <w:style w:type="paragraph" w:styleId="TestStyle">
    <w:name w:val="TestStyle"/>
    <w:pPr>${pPrInner}</w:pPr>
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

function getSpacing(doc: Document) {
  return doc.getStylesManager().getStyle('TestStyle')?.getParagraphFormatting()?.spacing;
}

describe('Style w:spacing — beforeLines / afterLines (CJK-friendly line units)', () => {
  it('parses beforeLines', async () => {
    const buffer = await makeDocxWithStylePPr('<w:spacing w:beforeLines="100"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getSpacing(doc)?.beforeLines).toBe(100);
    doc.dispose();
  });

  it('parses afterLines', async () => {
    const buffer = await makeDocxWithStylePPr('<w:spacing w:afterLines="150"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getSpacing(doc)?.afterLines).toBe(150);
    doc.dispose();
  });

  it('round-trips beforeLines + afterLines through load → save → load', async () => {
    const buffer1 = await makeDocxWithStylePPr(
      '<w:spacing w:beforeLines="100" w:afterLines="150"/>'
    );
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(getSpacing(doc1)?.beforeLines).toBe(100);
    expect(getSpacing(doc1)?.afterLines).toBe(150);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(getSpacing(doc2)?.beforeLines).toBe(100);
    expect(getSpacing(doc2)?.afterLines).toBe(150);
    doc2.dispose();
  });
});

describe('Style w:spacing — beforeAutospacing / afterAutospacing (ST_OnOff)', () => {
  it('parses beforeAutospacing="1" as true', async () => {
    const buffer = await makeDocxWithStylePPr('<w:spacing w:beforeAutospacing="1"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getSpacing(doc)?.beforeAutospacing).toBe(true);
    doc.dispose();
  });

  it('parses afterAutospacing="on" as true (ST_OnOff literal)', async () => {
    const buffer = await makeDocxWithStylePPr('<w:spacing w:afterAutospacing="on"/>');
    const doc = await Document.loadFromBuffer(buffer);
    expect(getSpacing(doc)?.afterAutospacing).toBe(true);
    doc.dispose();
  });

  it('round-trips beforeAutospacing through load → save → load', async () => {
    const buffer1 = await makeDocxWithStylePPr('<w:spacing w:beforeAutospacing="1"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(getSpacing(doc1)?.beforeAutospacing).toBe(true);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(getSpacing(doc2)?.beforeAutospacing).toBe(true);
    doc2.dispose();
  });

  it('round-trips afterAutospacing through load → save → load', async () => {
    const buffer1 = await makeDocxWithStylePPr('<w:spacing w:afterAutospacing="1"/>');
    const doc1 = await Document.loadFromBuffer(buffer1);
    expect(getSpacing(doc1)?.afterAutospacing).toBe(true);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    expect(getSpacing(doc2)?.afterAutospacing).toBe(true);
    doc2.dispose();
  });
});

describe('Style.toXML — generator emits all eight CT_Spacing attributes', () => {
  // Direct generator test — bypasses the styles.xml passthrough that
  // preserves raw XML when no styles are modified. This is the path exercised
  // when (a) creating a new document from scratch or (b) modifying a style
  // programmatically so `mergeStylesWithOriginal` regenerates it.
  it('emits beforeLines when set', () => {
    const style = Style.create({
      styleId: 'Test',
      name: 'Test',
      type: 'paragraph',
      paragraphFormatting: { spacing: { beforeLines: 100 } },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toMatch(/<w:spacing[^>]*w:beforeLines="100"/);
  });

  it('emits afterLines when set', () => {
    const style = Style.create({
      styleId: 'Test',
      name: 'Test',
      type: 'paragraph',
      paragraphFormatting: { spacing: { afterLines: 150 } },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toMatch(/<w:spacing[^>]*w:afterLines="150"/);
  });

  it('emits beforeAutospacing="1" when true', () => {
    const style = Style.create({
      styleId: 'Test',
      name: 'Test',
      type: 'paragraph',
      paragraphFormatting: { spacing: { beforeAutospacing: true } },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toMatch(/<w:spacing[^>]*w:beforeAutospacing="1"/);
  });

  it('emits afterAutospacing="1" when true', () => {
    const style = Style.create({
      styleId: 'Test',
      name: 'Test',
      type: 'paragraph',
      paragraphFormatting: { spacing: { afterAutospacing: true } },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toMatch(/<w:spacing[^>]*w:afterAutospacing="1"/);
  });

  it('emits beforeAutospacing="0" for explicit false (override of style-inherited true)', () => {
    const style = Style.create({
      styleId: 'Test',
      name: 'Test',
      type: 'paragraph',
      paragraphFormatting: { spacing: { beforeAutospacing: false } },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toMatch(/<w:spacing[^>]*w:beforeAutospacing="0"/);
  });

  it('emits all eight attributes together', () => {
    const style = Style.create({
      styleId: 'Test',
      name: 'Test',
      type: 'paragraph',
      paragraphFormatting: {
        spacing: {
          before: 240,
          beforeLines: 100,
          beforeAutospacing: false,
          after: 120,
          afterLines: 50,
          afterAutospacing: false,
          line: 360,
          lineRule: 'auto',
        },
      },
    });
    const xml = XMLBuilder.elementToString(style.toXML());
    expect(xml).toMatch(/w:before="240"/);
    expect(xml).toMatch(/w:beforeLines="100"/);
    expect(xml).toMatch(/w:beforeAutospacing="0"/);
    expect(xml).toMatch(/w:after="120"/);
    expect(xml).toMatch(/w:afterLines="50"/);
    expect(xml).toMatch(/w:afterAutospacing="0"/);
    expect(xml).toMatch(/w:line="360"/);
    expect(xml).toMatch(/w:lineRule="auto"/);
  });
});

describe('Style w:spacing — all eight attributes together (load-modify-save path)', () => {
  it('round-trips the full CT_Spacing attribute set', async () => {
    const buffer1 = await makeDocxWithStylePPr(
      '<w:spacing w:before="240" w:beforeLines="100" w:beforeAutospacing="0" w:after="120" w:afterLines="50" w:afterAutospacing="0" w:line="360" w:lineRule="auto"/>'
    );
    const doc1 = await Document.loadFromBuffer(buffer1);
    const s1 = getSpacing(doc1);
    expect(s1?.before).toBe(240);
    expect(s1?.beforeLines).toBe(100);
    expect(s1?.beforeAutospacing).toBe(false);
    expect(s1?.after).toBe(120);
    expect(s1?.afterLines).toBe(50);
    expect(s1?.afterAutospacing).toBe(false);
    expect(s1?.line).toBe(360);
    expect(s1?.lineRule).toBe('auto');

    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2);
    const s2 = getSpacing(doc2);
    expect(s2?.before).toBe(240);
    expect(s2?.beforeLines).toBe(100);
    expect(s2?.beforeAutospacing).toBe(false);
    expect(s2?.after).toBe(120);
    expect(s2?.afterLines).toBe(50);
    expect(s2?.afterAutospacing).toBe(false);
    expect(s2?.line).toBe(360);
    expect(s2?.lineRule).toBe('auto');
    doc2.dispose();
  });
});
