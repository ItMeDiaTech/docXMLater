/**
 * pPrChange previousProperties — CJK typography flags round-trip.
 *
 * `DocumentParser.parseParagraphProperties` reads `<w:pPrChange>` for
 * tracked paragraph-property revisions, but its `previousProperties`
 * block was missing parse lines for four CJK-typography CT_OnOff flags:
 *
 *   - `w:kinsoku`       (§17.3.1.16) — kinsoku line-break rules
 *   - `w:wordWrap`      (§17.3.1.45) — mid-word wrapping (actually parsed)
 *   - `w:overflowPunct` (§17.3.1.21) — punctuation hanging past margins
 *   - `w:topLinePunct`  (§17.3.1.43) — top-line punctuation compression
 *   - `w:suppressOverlap` (§17.3.1.34) — frame-overlap suppression
 *
 * The generator already emits all four correctly when they're present
 * in `previousProperties`, so the asymmetry was pure loss: load-then-save
 * a pPrChange containing any of these four flags and they're silently
 * dropped from the revision history.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithPPrChange(prevPPrInner: string): Promise<Buffer> {
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
        <w:pPrChange w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z">
          <w:pPr>${prevPPrInner}</w:pPr>
        </w:pPrChange>
      </w:pPr>
      <w:r><w:t>test</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('pPrChange previousProperties — CJK flags', () => {
  it('parses <w:kinsoku w:val="0"/> as previousProperties.kinsoku = false', async () => {
    const buffer = await makeDocxWithPPrChange('<w:kinsoku w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const prev = doc.getParagraphs()[0]!.getFormatting().pPrChange?.previousProperties;
    expect((prev as { kinsoku?: boolean })?.kinsoku).toBe(false);
    doc.dispose();
  });

  it('parses bare <w:kinsoku/> as previousProperties.kinsoku = true', async () => {
    const buffer = await makeDocxWithPPrChange('<w:kinsoku/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const prev = doc.getParagraphs()[0]!.getFormatting().pPrChange?.previousProperties;
    expect((prev as { kinsoku?: boolean })?.kinsoku).toBe(true);
    doc.dispose();
  });

  it('parses <w:overflowPunct/> as previousProperties.overflowPunct = true', async () => {
    const buffer = await makeDocxWithPPrChange('<w:overflowPunct/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const prev = doc.getParagraphs()[0]!.getFormatting().pPrChange?.previousProperties;
    expect((prev as { overflowPunct?: boolean })?.overflowPunct).toBe(true);
    doc.dispose();
  });

  it('parses <w:topLinePunct w:val="0"/> as previousProperties.topLinePunct = false', async () => {
    const buffer = await makeDocxWithPPrChange('<w:topLinePunct w:val="0"/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const prev = doc.getParagraphs()[0]!.getFormatting().pPrChange?.previousProperties;
    expect((prev as { topLinePunct?: boolean })?.topLinePunct).toBe(false);
    doc.dispose();
  });

  it('parses <w:suppressOverlap/> as previousProperties.suppressOverlap = true', async () => {
    const buffer = await makeDocxWithPPrChange('<w:suppressOverlap/>');
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const prev = doc.getParagraphs()[0]!.getFormatting().pPrChange?.previousProperties;
    expect((prev as { suppressOverlap?: boolean })?.suppressOverlap).toBe(true);
    doc.dispose();
  });

  it('round-trips all four flags through Document save/load', async () => {
    const buffer = await makeDocxWithPPrChange(
      `<w:kinsoku/>
       <w:overflowPunct w:val="0"/>
       <w:topLinePunct/>
       <w:suppressOverlap w:val="0"/>`
    );
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const rebuffered = await doc.toBuffer();
    doc.dispose();
    const reloaded = await Document.loadFromBuffer(rebuffered, { revisionHandling: 'preserve' });
    const prev = reloaded.getParagraphs()[0]!.getFormatting().pPrChange?.previousProperties as {
      kinsoku?: boolean;
      overflowPunct?: boolean;
      topLinePunct?: boolean;
      suppressOverlap?: boolean;
    };
    expect(prev?.kinsoku).toBe(true);
    expect(prev?.overflowPunct).toBe(false);
    expect(prev?.topLinePunct).toBe(true);
    expect(prev?.suppressOverlap).toBe(false);
    reloaded.dispose();
  });
});
