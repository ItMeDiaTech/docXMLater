/**
 * Tests for body-level w:bookmarkStart parsing in DocumentParser
 *
 * Issue: Body-level bookmarkStart elements (between paragraphs/tables) were
 * silently dropped, while bookmarkEnd elements at body level were preserved.
 * This caused orphaned bookmarkEnds without matching bookmarkStarts.
 *
 * Fix: Added extractBodyLevelBookmarkStarts() to capture body-level starts
 * and attach them to the NEXT parsed element (mirroring bookmarkEnd → PREVIOUS).
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

/**
 * Creates a minimal DOCX buffer with custom document.xml content.
 */
async function createMinimalDocx(documentXml: string): Promise<Buffer> {
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
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
  );

  zipHandler.addFile('word/document.xml', documentXml);

  return await zipHandler.toBuffer();
}

describe('Body-Level Bookmark Round-Trip', () => {
  it('should preserve body-level bookmarkStart elements', async () => {
    // Body-level bookmarkStart between two paragraphs
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Before bookmark</w:t></w:r>
    </w:p>
    <w:bookmarkStart w:id="24" w:name="_TestBookmark"/>
    <w:p>
      <w:r><w:t>After bookmark start</w:t></w:r>
      <w:bookmarkEnd w:id="24"/>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createMinimalDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);

    const paragraphs = doc.getParagraphs();
    expect(paragraphs.length).toBe(2);

    // The bookmarkStart should be attached to the NEXT paragraph (index 1)
    const secondPara = paragraphs[1]!;
    const starts = secondPara.getBookmarksStart();
    expect(starts.length).toBe(1);
    expect(starts[0]!.getName()).toBe('_TestBookmark');
    expect(starts[0]!.getId()).toBe(24);

    // The bookmarkEnd should be in the same paragraph (parsed from within w:p)
    const ends = secondPara.getBookmarksEnd();
    expect(ends.length).toBe(1);
    expect(ends[0]!.getId()).toBe(24);

    // Round-trip: save and verify output XML contains both start and end
    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml') || '';

    expect(outputXml).toContain('w:bookmarkStart');
    expect(outputXml).toContain('w:bookmarkEnd');
    expect(outputXml).toContain('_TestBookmark');
  });

  it('should handle multiple clustered bookmarkStarts between elements', async () => {
    // Simulates the Option_1.docx pattern: multiple bookmarkStarts in a cluster
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Paragraph before cluster</w:t></w:r>
    </w:p>
    <w:bookmarkStart w:id="24" w:name="_PAR_Process"/>
    <w:bookmarkStart w:id="25" w:name="_Next_Day"/>
    <w:bookmarkStart w:id="26" w:name="_Heading_3"/>
    <w:p>
      <w:r><w:t>Paragraph after cluster</w:t></w:r>
    </w:p>
    <w:bookmarkEnd w:id="24"/>
    <w:bookmarkEnd w:id="25"/>
    <w:bookmarkEnd w:id="26"/>
    <w:p>
      <w:r><w:t>Final paragraph</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createMinimalDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);

    const paragraphs = doc.getParagraphs();
    expect(paragraphs.length).toBe(3);

    // All 3 bookmarkStarts should be attached to the second paragraph
    const secondPara = paragraphs[1]!;
    const starts = secondPara.getBookmarksStart();
    expect(starts.length).toBe(3);

    const names = starts.map((b) => b.getName()).sort();
    expect(names).toEqual(['_Heading_3', '_Next_Day', '_PAR_Process']);

    // The bookmarkEnds should be attached to the second paragraph (body-level → previous)
    const ends = secondPara.getBookmarksEnd();
    expect(ends.length).toBe(3);

    doc.dispose();
  });

  it('should handle mixed bookmarkStarts and bookmarkEnds between elements', async () => {
    // bookmarkEnds from a previous bookmark + bookmarkStarts for the next region
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id="10" w:name="_Earlier"/>
      <w:r><w:t>First paragraph</w:t></w:r>
    </w:p>
    <w:bookmarkEnd w:id="10"/>
    <w:bookmarkStart w:id="20" w:name="_Later"/>
    <w:p>
      <w:r><w:t>Second paragraph</w:t></w:r>
      <w:bookmarkEnd w:id="20"/>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createMinimalDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);

    const paragraphs = doc.getParagraphs();
    expect(paragraphs.length).toBe(2);

    // bookmarkEnd id=10 → attached to PREVIOUS (first paragraph)
    const firstEnds = paragraphs[0]!.getBookmarksEnd();
    expect(firstEnds.length).toBe(1);
    expect(firstEnds[0]!.getId()).toBe(10);

    // bookmarkStart id=20 → attached to NEXT (second paragraph)
    const secondStarts = paragraphs[1]!.getBookmarksStart();
    expect(secondStarts.length).toBe(1);
    expect(secondStarts[0]!.getName()).toBe('_Later');

    doc.dispose();
  });

  it('should handle bookmarkStarts before the first element', async () => {
    // bookmarkStart at the very beginning of the body
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:bookmarkStart w:id="1" w:name="_DocStart"/>
    <w:p>
      <w:r><w:t>First paragraph</w:t></w:r>
      <w:bookmarkEnd w:id="1"/>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createMinimalDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);

    const paragraphs = doc.getParagraphs();
    expect(paragraphs.length).toBe(1);

    // bookmarkStart should be attached to the first (and only) paragraph
    const starts = paragraphs[0]!.getBookmarksStart();
    expect(starts.length).toBe(1);
    expect(starts[0]!.getName()).toBe('_DocStart');

    doc.dispose();
  });

  it('should count equal bookmarkStarts and bookmarkEnds after round-trip', async () => {
    // The full pattern from Option_1.docx: 7 bookmarkStarts and 7 bookmarkEnds at body level
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Before bookmarks</w:t></w:r>
    </w:p>
    <w:bookmarkStart w:id="24" w:name="_Bookmark_A"/>
    <w:bookmarkStart w:id="25" w:name="_Bookmark_B"/>
    <w:bookmarkStart w:id="26" w:name="_Bookmark_C"/>
    <w:bookmarkStart w:id="27" w:name="_Bookmark_D"/>
    <w:bookmarkStart w:id="28" w:name="_Bookmark_E"/>
    <w:bookmarkStart w:id="29" w:name="_Bookmark_F"/>
    <w:bookmarkStart w:id="30" w:name="_Bookmark_G"/>
    <w:p>
      <w:r><w:t>Middle paragraph</w:t></w:r>
    </w:p>
    <w:bookmarkEnd w:id="24"/>
    <w:bookmarkEnd w:id="25"/>
    <w:bookmarkEnd w:id="26"/>
    <w:bookmarkEnd w:id="27"/>
    <w:bookmarkEnd w:id="28"/>
    <w:bookmarkEnd w:id="29"/>
    <w:bookmarkEnd w:id="30"/>
    <w:p>
      <w:r><w:t>After bookmarks</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createMinimalDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);

    // Save and check output
    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml') || '';

    // Count bookmarkStart and bookmarkEnd occurrences in output
    const startMatches = outputXml.match(/w:bookmarkStart/g) || [];
    const endMatches = outputXml.match(/w:bookmarkEnd/g) || [];

    // Every bookmarkStart should have a matching bookmarkEnd
    expect(startMatches.length).toBe(7);
    expect(endMatches.length).toBe(7);

    // Verify no orphaned ends (each end ID has a matching start ID)
    for (let id = 24; id <= 30; id++) {
      expect(outputXml).toContain(`w:name="_Bookmark_${String.fromCharCode(65 + id - 24)}"`);
      expect(outputXml).toContain(`<w:bookmarkEnd w:id="${id}"`);
    }
  });

  it('should handle body-level bookmarkStarts before a table', async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Before table</w:t></w:r>
    </w:p>
    <w:bookmarkStart w:id="5" w:name="_TableBookmark"/>
    <w:tbl>
      <w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:p><w:r><w:t>Cell content</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:bookmarkEnd w:id="5"/>
    <w:p>
      <w:r><w:t>After table</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const buffer = await createMinimalDocx(documentXml);
    const doc = await Document.loadFromBuffer(buffer);

    // The bookmarkStart should be attached to the table's first cell first paragraph
    const tables = doc.getTables();
    expect(tables.length).toBe(1);

    const firstCell = tables[0]!.getRow(0)!.getCell(0)!;
    const cellParas = firstCell.getParagraphs();
    expect(cellParas.length).toBeGreaterThan(0);

    const starts = cellParas[0]!.getBookmarksStart();
    expect(starts.length).toBe(1);
    expect(starts[0]!.getName()).toBe('_TableBookmark');

    // Round-trip verification
    const outputBuffer = await doc.toBuffer();
    doc.dispose();

    const outputZip = new ZipHandler();
    await outputZip.loadFromBuffer(outputBuffer);
    const outputXml = outputZip.getFileAsString('word/document.xml') || '';

    expect(outputXml).toContain('_TableBookmark');
    expect(outputXml).toContain('<w:bookmarkEnd w:id="5"');
  });
});
