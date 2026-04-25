/**
 * TabAlignment — ST_TabJc `start` / `end` values.
 *
 * Per ECMA-376 Part 1 §17.18.94 ST_TabJc has 9 values (clear, start,
 * center, end, decimal, bar, num, left, right). The `TabAlignment`
 * type was declared with only 7 — missing the bidi-aware `start` and
 * `end` variants — so a paragraph with `<w:tab w:val="start" w:pos="720"/>`
 * (the modern default emitted by bidi-aware authoring tools) would fail
 * TypeScript assignment checks even though the parser passes the raw
 * string through.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithTabs(tabsXml: string): Promise<Buffer> {
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
    <w:p>
      <w:pPr><w:tabs>${tabsXml}</w:tabs></w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('TabAlignment — ST_TabJc `start` / `end` (§17.18.94)', () => {
  it('parses <w:tab w:val="start" w:pos="720"/> and preserves alignment as "start"', async () => {
    const buffer = await makeDocxWithTabs('<w:tab w:val="start" w:pos="720"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs?.[0]?.val).toBe('start');
    expect(tabs?.[0]?.position).toBe(720);
    doc.dispose();
  });

  it('parses <w:tab w:val="end" w:pos="9360"/> and preserves alignment as "end"', async () => {
    const buffer = await makeDocxWithTabs('<w:tab w:val="end" w:pos="9360"/>');
    const doc = await Document.loadFromBuffer(buffer);
    const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
    expect(tabs?.[0]?.val).toBe('end');
    expect(tabs?.[0]?.position).toBe(9360);
    doc.dispose();
  });

  it('parses the remaining 7 ST_TabJc values without loss', async () => {
    const values = ['clear', 'left', 'center', 'right', 'decimal', 'bar', 'num'];
    for (const val of values) {
      const buffer = await makeDocxWithTabs(`<w:tab w:val="${val}" w:pos="720"/>`);
      const doc = await Document.loadFromBuffer(buffer);
      const tabs = doc.getParagraphs()[0]!.getFormatting().tabs;
      expect(tabs?.[0]?.val).toBe(val);
      doc.dispose();
    }
  });
});
