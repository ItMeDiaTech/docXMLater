/**
 * Raw-XML revision acceptor — structural integrity when tracked-deleted
 * rows are filtered out of a table.
 *
 * Two failure modes guarded against:
 * 1. Filtering `w:tr` entries without rebuilding `_orderedChildren` left
 *    stale {type,index} entries; the serializer maps entries onto the
 *    shrunken row array by index, so inter-row siblings (e.g. a
 *    `w:bookmarkEnd` between rows, legal per ECMA-376 CT_Tbl) shifted
 *    relative to surviving rows and trailing rows were silently dropped.
 * 2. When every row carried a `trPr` `w:del` marker, only the `w:tr` key
 *    was deleted, emitting a row-less `<w:tbl>` — invalid per ECMA-376
 *    §17.4.38 (at least one `w:tr` required). Word's Accept All removes
 *    the entire table in this case, so the acceptor must too.
 */
import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';
import { acceptAllRevisions } from '../../src/processors/acceptRevisions';

function makeZip(documentXml: string): ZipHandler {
  const zip = new ZipHandler();
  zip.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );
  zip.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.addFile('word/document.xml', documentXml);
  return zip;
}

function wrapBody(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
${body}
    <w:p><w:r><w:t>after</w:t></w:r></w:p>
  </w:body>
</w:document>`;
}

const DELETED_ROW = `      <w:tr>
        <w:trPr><w:del w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:trPr>
        <w:tc><w:tcPr/><w:p><w:r><w:t>gone</w:t></w:r></w:p></w:tc>
      </w:tr>`;

describe('acceptAllRevisions (raw-XML DOM path) — row filtering integrity', () => {
  it('keeps an inter-row bookmarkEnd anchored when a preceding row is deleted', async () => {
    const zip = makeZip(
      wrapBody(`    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
${DELETED_ROW}
      <w:bookmarkEnd w:id="5"/>
      <w:tr><w:trPr/><w:tc><w:tcPr/><w:p><w:r><w:t>row1</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:trPr/><w:tc><w:tcPr/><w:p><w:r><w:t>row2</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>`)
    );
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/document.xml')!;

    expect(after).not.toContain('gone');
    // Both surviving rows must be emitted (stale entries dropped the tail).
    expect(after).toContain('row1');
    expect(after).toContain('row2');
    // The marker sat before the first surviving row; it must stay there.
    const markerIdx = after.indexOf('<w:bookmarkEnd');
    const firstRowIdx = after.indexOf('<w:tr');
    expect(markerIdx).toBeGreaterThan(-1);
    expect(firstRowIdx).toBeGreaterThan(-1);
    expect(markerIdx).toBeLessThan(firstRowIdx);
  });

  it('removes the entire table when every row (array form) is tracked-deleted', async () => {
    const zip = makeZip(
      wrapBody(`    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
${DELETED_ROW}
      <w:tr>
        <w:trPr><w:del w:id="2" w:author="A" w:date="2026-01-15T10:00:00Z"/></w:trPr>
        <w:tc><w:tcPr/><w:p><w:r><w:t>gone too</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>`)
    );
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/document.xml')!;

    // A row-less <w:tbl> (only tblPr/tblGrid) is invalid OOXML; the whole
    // table must go, leaving surrounding content intact.
    expect(after).not.toMatch(/<w:tbl[ >/]/);
    expect(after).not.toContain('<w:tblGrid');
    expect(after).toContain('after');
  });

  it('removes the entire table when its single row is tracked-deleted', async () => {
    const zip = makeZip(
      wrapBody(`    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
${DELETED_ROW}
    </w:tbl>`)
    );
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/document.xml')!;

    expect(after).not.toMatch(/<w:tbl[ >/]/);
    expect(after).toContain('after');
  });

  it('fully deleted table does not reach the saved document (end-to-end)', async () => {
    const zip = makeZip(
      wrapBody(`    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
${DELETED_ROW}
    </w:tbl>`)
    );
    const buffer = await zip.toBuffer();
    // Default revisionHandling: 'accept' invokes the raw-XML acceptor.
    const doc = await Document.loadFromBuffer(buffer);
    try {
      expect(doc.getBodyElements().some((el) => el instanceof Table)).toBe(false);
      const saved = await doc.toBuffer();
      const savedZip = new ZipHandler();
      await savedZip.loadFromBuffer(saved);
      const savedXml = savedZip.getFileAsString('word/document.xml')!;
      expect(savedXml).not.toMatch(/<w:tbl[ >/]/);
      expect(savedXml).toContain('after');
    } finally {
      doc.dispose();
    }
  });
});
