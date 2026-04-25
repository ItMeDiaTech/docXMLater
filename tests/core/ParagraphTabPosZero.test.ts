/**
 * Paragraph `<w:tab>` parser — `w:pos="0"` falsy-check bug.
 *
 * Per ECMA-376 Part 1 §17.3.1.38, `<w:tab>` inside `<w:tabs>` has:
 *
 *   w:val    — ST_TabJc (alignment) — optional
 *   w:pos    — ST_SignedTwipsMeasure (position) — REQUIRED
 *   w:leader — ST_TabTlc (leader char) — optional
 *
 * `w:pos` can legitimately be:
 *   - `0` (tab at left margin / start)
 *   - negative (tab extending into the margin)
 *   - positive (standard case)
 *
 * Bug this suite guards against:
 *   - DocumentParser.parseParagraphFromObject contained:
 *         if (tabObj['@_w:pos']) tab.position = parseInt(...);
 *     XMLParser with parseAttributeValue:true coerces "0" → number 0,
 *     which is falsy — so the position was never set, and the tab was
 *     then dropped entirely by the downstream `tab.position !== undefined`
 *     guard. Any paragraph with a tab at position 0 silently lost that
 *     tab on load. (The `w:pPrChange` tab parser at the same file
 *     already used a `!== undefined` check, so tracked "previous" tabs
 *     round-tripped correctly — making the asymmetry especially subtle.)
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithParagraphTabs(tabsXml: string): Promise<Buffer> {
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
    <w:p>
      <w:pPr>
        <w:tabs>${tabsXml}</w:tabs>
      </w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('Paragraph tab parser — w:pos="0" preserved', () => {
  it('parses a single tab at w:pos="0"', async () => {
    const buffer = await makeDocxWithParagraphTabs('<w:tab w:val="left" w:pos="0"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs).toHaveLength(1);
    expect(tabs![0]!.position).toBe(0);
    expect(tabs![0]!.val).toBe('left');
    doc.dispose();
  });

  it('parses multiple tabs including one at w:pos="0"', async () => {
    const buffer = await makeDocxWithParagraphTabs(
      '<w:tab w:val="left" w:pos="0"/><w:tab w:val="center" w:pos="2880"/><w:tab w:val="right" w:pos="5760"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs).toHaveLength(3);
    expect(tabs![0]!.position).toBe(0);
    expect(tabs![1]!.position).toBe(2880);
    expect(tabs![2]!.position).toBe(5760);
    doc.dispose();
  });

  it('parses a tab at negative position (extends into margin)', async () => {
    const buffer = await makeDocxWithParagraphTabs('<w:tab w:val="left" w:pos="-360"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs).toHaveLength(1);
    expect(tabs![0]!.position).toBe(-360);
    doc.dispose();
  });

  it('parses a tab with w:val and w:leader at position 0', async () => {
    const buffer = await makeDocxWithParagraphTabs(
      '<w:tab w:val="decimal" w:pos="0" w:leader="dot"/>'
    );
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs).toHaveLength(1);
    expect(tabs![0]!.position).toBe(0);
    expect(tabs![0]!.val).toBe('decimal');
    expect(tabs![0]!.leader).toBe('dot');
    doc.dispose();
  });

  it('still parses tabs at positive positions correctly (regression check)', async () => {
    const buffer = await makeDocxWithParagraphTabs('<w:tab w:val="left" w:pos="720"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs).toHaveLength(1);
    expect(tabs![0]!.position).toBe(720);
    doc.dispose();
  });

  it('drops tabs with completely missing w:pos (schema-required attribute)', async () => {
    // Per §17.3.1.38 w:pos is REQUIRED, so a tab without pos is malformed.
    // Parser correctly drops these (this is an existing behaviour — locked here
    // so the pos=0 fix doesn't accidentally start accepting pos-less tabs).
    const buffer = await makeDocxWithParagraphTabs('<w:tab w:val="left"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs === undefined || tabs.length === 0).toBe(true);
    doc.dispose();
  });
});
