/**
 * w:pPrChange tracked-change previousProperties — CJK character-unit
 * indentation round-trip.
 *
 * Iteration 21 added `w:leftChars` / `w:rightChars` / `w:firstLineChars` /
 * `w:hangingChars` support to the main-path paragraph `w:ind` parser and
 * generator. The pPrChange sibling parser at DocumentParser.ts:2608 and
 * the pPrChange generator at Paragraph.ts:3740 both used the exact same
 * four-attribute twips-only code shape, so tracked "previous
 * properties" for any CJK-authored document with character-unit indents
 * STILL silently lost those attributes on load → save, even after the
 * iteration-21 fix.
 *
 * Example the previous-state half of a tracked indent change:
 *
 *   <w:pPrChange w:id="1" w:author="…" w:date="…">
 *     <w:pPr>
 *       <w:ind w:leftChars="400" w:firstLineChars="100"/>
 *     </w:pPr>
 *   </w:pPrChange>
 *
 * Without this fix, reading that `<w:ind>` dropped `leftChars` and
 * `firstLineChars` — the tracked "previous" indentation state
 * degenerated into an empty `<w:ind/>` on save, breaking Word's ability
 * to accurately render what the paragraph looked like before the change.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithPPrChangeInd(indXml: string): Promise<Buffer> {
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
        <w:pPrChange w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z">
          <w:pPr>${indXml}</w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

// The pPrChange parser stores previousProperties as Record<string, any>, so
// type-unsafe access is needed. We retrieve through getFormatting's pPrChange
// handle and inspect the indentation sub-object.
function getPrevIndent(doc: Document): Record<string, unknown> | undefined {
  const change = doc.getParagraphs()[0]!.formatting.pPrChange;
  const props = change?.previousProperties as Record<string, unknown> | undefined;
  return props?.indentation as Record<string, unknown> | undefined;
}

describe('w:pPrChange — CJK character-unit indentation parsing', () => {
  it('parses w:firstLineChars in previousProperties', async () => {
    const buffer = await makeDocxWithPPrChangeInd('<w:ind w:firstLineChars="200"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const ind = getPrevIndent(doc);
    expect(ind?.firstLineChars).toBe(200);
    doc.dispose();
  });

  it('parses w:hangingChars in previousProperties', async () => {
    const buffer = await makeDocxWithPPrChangeInd('<w:ind w:hangingChars="150"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const ind = getPrevIndent(doc);
    expect(ind?.hangingChars).toBe(150);
    doc.dispose();
  });

  it('parses w:leftChars in previousProperties', async () => {
    const buffer = await makeDocxWithPPrChangeInd('<w:ind w:leftChars="400"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const ind = getPrevIndent(doc);
    expect(ind?.leftChars).toBe(400);
    doc.dispose();
  });

  it('parses w:rightChars in previousProperties', async () => {
    const buffer = await makeDocxWithPPrChangeInd('<w:ind w:rightChars="300"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const ind = getPrevIndent(doc);
    expect(ind?.rightChars).toBe(300);
    doc.dispose();
  });

  it('parses w:startChars (bidi-aware) as leftChars in previousProperties', async () => {
    const buffer = await makeDocxWithPPrChangeInd('<w:ind w:startChars="400"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const ind = getPrevIndent(doc);
    expect(ind?.leftChars).toBe(400);
    doc.dispose();
  });

  it('parses w:endChars (bidi-aware) as rightChars in previousProperties', async () => {
    const buffer = await makeDocxWithPPrChangeInd('<w:ind w:endChars="300"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const ind = getPrevIndent(doc);
    expect(ind?.rightChars).toBe(300);
    doc.dispose();
  });

  it('parses firstLineChars=0 in previousProperties (number-0 trap guard)', async () => {
    const buffer = await makeDocxWithPPrChangeInd('<w:ind w:firstLineChars="0"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const ind = getPrevIndent(doc);
    expect(ind?.firstLineChars).toBe(0);
    doc.dispose();
  });
});

describe('w:pPrChange — CJK character-unit indentation full round-trip', () => {
  it('round-trips firstLineChars through load → save → load', async () => {
    const buffer1 = await makeDocxWithPPrChangeInd('<w:ind w:firstLineChars="200"/>');
    const doc1 = await Document.loadFromBuffer(buffer1, { revisionHandling: 'preserve' });
    expect(getPrevIndent(doc1)?.firstLineChars).toBe(200);
    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2, { revisionHandling: 'preserve' });
    expect(getPrevIndent(doc2)?.firstLineChars).toBe(200);
    doc2.dispose();
  });

  it('round-trips left/rightChars together', async () => {
    const buffer1 = await makeDocxWithPPrChangeInd(
      '<w:ind w:leftChars="400" w:rightChars="200" w:firstLineChars="100"/>'
    );
    const doc1 = await Document.loadFromBuffer(buffer1, { revisionHandling: 'preserve' });
    const ind1 = getPrevIndent(doc1);
    expect(ind1?.leftChars).toBe(400);
    expect(ind1?.rightChars).toBe(200);
    expect(ind1?.firstLineChars).toBe(100);

    const buffer2 = await doc1.toBuffer();
    doc1.dispose();

    const doc2 = await Document.loadFromBuffer(buffer2, { revisionHandling: 'preserve' });
    const ind2 = getPrevIndent(doc2);
    expect(ind2?.leftChars).toBe(400);
    expect(ind2?.rightChars).toBe(200);
    expect(ind2?.firstLineChars).toBe(100);
    doc2.dispose();
  });
});
