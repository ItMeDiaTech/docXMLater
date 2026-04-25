/**
 * ParagraphAlignment — ST_Jc value `numTab` round-trip.
 *
 * Per ECMA-376 Part 1 §17.18.44 ST_Jc enum, `numTab` is one of the 13
 * paragraph-alignment values; the `ParagraphAlignment` type omitted it
 * (11 values only) and the type guard `isParagraphAlignment` also
 * rejected it. Rare in practice but still spec-valid — when used by a
 * numbered paragraph, it aligns to the numbering tab stop.
 */

import { Document } from '../../src/core/Document';
import { isParagraphAlignment } from '../../src/elements/CommonTypes';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithJc(jcVal: string): Promise<Buffer> {
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
      <w:pPr><w:jc w:val="${jcVal}"/></w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('ParagraphAlignment — ST_Jc `numTab` (§17.18.44)', () => {
  it('isParagraphAlignment accepts "numTab" as valid', () => {
    expect(isParagraphAlignment('numTab')).toBe(true);
  });

  it('isParagraphAlignment still rejects unknown values', () => {
    expect(isParagraphAlignment('bogus')).toBe(false);
  });

  it('parses <w:jc w:val="numTab"/> and preserves it on the paragraph', async () => {
    const buffer = await makeDocxWithJc('numTab');
    const doc = await Document.loadFromBuffer(buffer);
    expect(doc.getParagraphs()[0]!.getFormatting().alignment).toBe('numTab');
    doc.dispose();
  });
});
