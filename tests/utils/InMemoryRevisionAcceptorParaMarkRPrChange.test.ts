/**
 * InMemoryRevisionAcceptor — paragraphMarkRunPropertiesChange acceptance.
 *
 * Per ECMA-376 Part 1 §17.3.1.30 (CT_ParaRPrChange) the
 * `paragraphMarkRunPropertiesChange` field tracks the previous rPr of
 * the paragraph-mark glyph under a property-change revision. Accepting
 * `propertyChanges` must clear this field (the current formatting is
 * already the accepted state — only the "previous" snapshot is stale).
 *
 * Bug guarded against: after iteration 33 added this field to
 * `ParagraphFormatting`, the InMemoryRevisionAcceptor was not extended
 * to clear it. `acceptRevisionsInMemory({ acceptPropertyChanges: true })`
 * would silently leave the paragraph-mark rPrChange in place, so the
 * re-saved document kept advertising a stale tracked change.
 */

import { Document } from '../../src/core/Document';
import { acceptRevisionsInMemory } from '../../src/processors/InMemoryRevisionAcceptor';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithParaMarkRPrChange(): Promise<Buffer> {
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
      <w:pPr>
        <w:rPr>
          <w:b/>
          <w:rPrChange w:id="7" w:author="Jane" w:date="2026-01-15T10:00:00Z">
            <w:rPr><w:i/></w:rPr>
          </w:rPrChange>
        </w:rPr>
      </w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('acceptRevisionsInMemory — paragraphMarkRunPropertiesChange', () => {
  it('clears paragraphMarkRunPropertiesChange when acceptPropertyChanges is true', async () => {
    const buffer = await makeDocxWithParaMarkRPrChange();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const paraBefore = doc.getParagraphs()[0]!;
    expect(paraBefore.getFormatting().paragraphMarkRunPropertiesChange).toBeDefined();

    const result = acceptRevisionsInMemory(doc, {
      acceptInsertions: false,
      acceptDeletions: false,
      acceptMoves: false,
      acceptPropertyChanges: true,
    });

    const paraAfter = doc.getParagraphs()[0]!;
    expect(paraAfter.getFormatting().paragraphMarkRunPropertiesChange).toBeUndefined();
    expect(result.propertyChangesAccepted).toBeGreaterThanOrEqual(1);
    doc.dispose();
  });

  it('leaves paragraphMarkRunPropertiesChange intact when acceptPropertyChanges is false', async () => {
    const buffer = await makeDocxWithParaMarkRPrChange();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    acceptRevisionsInMemory(doc, {
      acceptInsertions: true,
      acceptDeletions: true,
      acceptMoves: true,
      acceptPropertyChanges: false,
    });

    const para = doc.getParagraphs()[0]!;
    expect(para.getFormatting().paragraphMarkRunPropertiesChange).toBeDefined();
    doc.dispose();
  });

  it('preserves current paragraph-mark formatting when clearing rPrChange', async () => {
    const buffer = await makeDocxWithParaMarkRPrChange();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const paraBefore = doc.getParagraphs()[0]!;
    expect(paraBefore.getFormatting().paragraphMarkRunProperties?.bold).toBe(true);

    acceptRevisionsInMemory(doc, {
      acceptInsertions: false,
      acceptDeletions: false,
      acceptMoves: false,
      acceptPropertyChanges: true,
    });

    const paraAfter = doc.getParagraphs()[0]!;
    // Current formatting (bold from the base rPr) is retained — only the
    // "previous" snapshot inside rPrChange is discarded.
    expect(paraAfter.getFormatting().paragraphMarkRunProperties?.bold).toBe(true);
    expect(paraAfter.getFormatting().paragraphMarkRunPropertiesChange).toBeUndefined();
    doc.dispose();
  });
});
