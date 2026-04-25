/**
 * Paragraph mark `<w:rPr>` — rPrChange tracked revision round-trip.
 *
 * Per ECMA-376 Part 1 §17.3.1.30 (CT_ParaRPrChange), the `<w:rPrChange>`
 * element may appear as the LAST child of a paragraph mark's `<w:rPr>` to
 * record the previous run-property set of the paragraph mark glyph
 * (pilcrow) before a tracked formatting change:
 *
 *   <w:pPr>
 *     <w:rPr>
 *       <w:ins .../>                    <!-- optional EG_ParaRPrTrackChanges -->
 *       <w:b/>                          <!-- EG_RPrBase — current formatting -->
 *       <w:rPrChange w:id="1" w:author="X" w:date="2026-01-15T10:00:00Z">
 *         <w:rPr><w:i/></w:rPr>         <!-- previous formatting -->
 *       </w:rPrChange>
 *     </w:rPr>
 *   </w:pPr>
 *
 * Bug guarded against: the paragraph-mark parser calls
 * `parseRunPropertiesFromObject` which correctly parses `<w:rPrChange>`
 * into `tempRun.propertyChangeRevision`, but the subsequent
 * `paragraph.setParagraphMarkFormatting(tempRun.getFormatting())` transfers
 * only the RunFormatting fields — `propertyChangeRevision` is a separate
 * field on Run that is silently dropped. Result: the paragraph mark's
 * tracked rPrChange is lost on load → save.
 */

import { Document } from '../../src/core/Document';
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
            <w:rPr>
              <w:i/>
            </w:rPr>
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

describe('Paragraph mark rPr — rPrChange (§17.3.1.30) round-trip', () => {
  it('parses <w:rPrChange> under <w:pPr><w:rPr> as paragraphMarkRunPropertiesChange', async () => {
    const buffer = await makeDocxWithParaMarkRPrChange();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const para = doc.getParagraphs()[0]!;
    const fmt = para.getFormatting();
    expect(fmt.paragraphMarkRunPropertiesChange).toBeDefined();
    expect(fmt.paragraphMarkRunPropertiesChange?.id).toBe(7);
    expect(fmt.paragraphMarkRunPropertiesChange?.author).toBe('Jane');
    // Previous formatting was italic (w:i).
    expect(fmt.paragraphMarkRunPropertiesChange?.previousProperties.italic).toBe(true);
    doc.dispose();
  });

  it('round-trips paragraph-mark rPrChange through Document save/load', async () => {
    const buffer = await makeDocxWithParaMarkRPrChange();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const rebuffered = await doc.toBuffer();
    doc.dispose();

    const reloaded = await Document.loadFromBuffer(rebuffered, { revisionHandling: 'preserve' });
    const fmt = reloaded.getParagraphs()[0]!.getFormatting();
    expect(fmt.paragraphMarkRunPropertiesChange).toBeDefined();
    expect(fmt.paragraphMarkRunPropertiesChange?.author).toBe('Jane');
    expect(fmt.paragraphMarkRunPropertiesChange?.previousProperties.italic).toBe(true);
    reloaded.dispose();
  });

  it('emits <w:rPrChange> inside the paragraph-mark <w:rPr> after re-save', async () => {
    const buffer = await makeDocxWithParaMarkRPrChange();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    // Force regeneration of document.xml (bypass raw-XML passthrough):
    // touching the paragraph mark formatting toggles the dirty flag.
    const para = doc.getParagraphs()[0]!;
    para.setParagraphMarkFormatting({
      ...para.getFormatting().paragraphMarkRunProperties,
      bold: true,
    });
    const rebuffered = await doc.toBuffer();
    doc.dispose();

    // Extract document.xml from the saved buffer.
    const zh = new ZipHandler();
    await zh.loadFromBuffer(rebuffered);
    const docXml = zh.getFileAsString('word/document.xml') ?? '';

    const pPrStart = docXml.indexOf('<w:pPr>');
    const pPrEnd = docXml.indexOf('</w:pPr>');
    const pPrBlock = pPrStart >= 0 && pPrEnd >= 0 ? docXml.substring(pPrStart, pPrEnd) : docXml;
    expect(pPrBlock).toContain('<w:rPrChange');
    const rPrChangeIdx = pPrBlock.indexOf('<w:rPrChange');
    const afterChange = pPrBlock.substring(rPrChangeIdx);
    expect(afterChange).toMatch(/<w:i(?:\/>|\s+w:val="1")/);
  });
});
