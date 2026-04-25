/**
 * CT_FramePr — ST_YAlign `inline` value round-trip.
 *
 * Per ECMA-376 Part 1 §17.18.110 ST_YAlign has six values:
 * `bottom`, `center`, `inline`, `inside`, `outside`, `top`.
 *
 * The `FrameProperties.yAlign` type was declared with only 5 values —
 * missing `inline`, which specifies that the frame is anchored in line
 * with the surrounding text rather than vertically offset. A paragraph
 * style using `<w:framePr w:yAlign="inline" .../>` would fail TypeScript
 * assignment even though the parser at DocumentParser.ts:~2491 passes
 * the raw string through unfiltered.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithFramePr(framePrAttrs: string): Promise<Buffer> {
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
      <w:pPr><w:framePr ${framePrAttrs}/></w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('CT_FramePr — ST_YAlign `inline` (§17.18.110)', () => {
  it('parses w:yAlign="inline" and preserves it on the paragraph framePr', async () => {
    const buffer = await makeDocxWithFramePr('w:yAlign="inline"');
    const doc = await Document.loadFromBuffer(buffer);
    const frame = doc.getParagraphs()[0]!.getFormatting().framePr;
    expect(frame?.yAlign).toBe('inline');
    doc.dispose();
  });

  it('parses all six ST_YAlign values without loss', async () => {
    const values = ['top', 'center', 'bottom', 'inline', 'inside', 'outside'] as const;
    for (const val of values) {
      const buffer = await makeDocxWithFramePr(`w:yAlign="${val}"`);
      const doc = await Document.loadFromBuffer(buffer);
      const frame = doc.getParagraphs()[0]!.getFormatting().framePr;
      expect(frame?.yAlign).toBe(val);
      doc.dispose();
    }
  });
});
