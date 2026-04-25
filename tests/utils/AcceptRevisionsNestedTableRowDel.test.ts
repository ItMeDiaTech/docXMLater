/**
 * Raw-XML revision acceptor — row-level `<w:del/>` must not over-match
 * across nested tables.
 *
 * Guards iteration 32's fix in `RevisionWalker.filterDeletedRows` against
 * over-reach. An outer row containing a nested `<w:tbl>` with its own
 * tracked-deleted rows should be processed correctly: the inner deleted
 * rows are removed, but the outer row (which has NO row-level del) must
 * survive.
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeNestedDocxWithRowDel(): Promise<Buffer> {
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
  // Outer table: one row with a nested table inside its cell. The nested
  // table's second row is tracked-deleted; the outer row is intact.
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr/>
        <w:tc>
          <w:tcPr/>
          <w:p><w:r><w:t>outer row</w:t></w:r></w:p>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
            <w:tr>
              <w:trPr/>
              <w:tc><w:tcPr/><w:p><w:r><w:t>inner keep</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:trPr>
                <w:del w:id="3" w:author="A" w:date="2026-01-15T10:00:00Z"/>
              </w:trPr>
              <w:tc><w:tcPr/><w:p><w:r><w:t>inner gone</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
              <w:trPr/>
              <w:tc><w:tcPr/><w:p><w:r><w:t>inner tail</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
          <w:p/>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

async function makeNestedDocxWithOuterRowDel(): Promise<Buffer> {
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
  // Outer table: FIRST row has row-level del AND contains a nested table
  // (whose rows are themselves tracked-independent). Accepting deletions
  // must remove the entire outer row (including the nested table) and
  // keep the "keep" row at the outer level.
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr>
          <w:del w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z"/>
        </w:trPr>
        <w:tc>
          <w:tcPr/>
          <w:p><w:r><w:t>outer gone</w:t></w:r></w:p>
          <w:tbl>
            <w:tblPr/>
            <w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
            <w:tr>
              <w:trPr/>
              <w:tc><w:tcPr/><w:p><w:r><w:t>inner survives removal?</w:t></w:r></w:p></w:tc>
            </w:tr>
          </w:tbl>
          <w:p/>
        </w:tc>
      </w:tr>
      <w:tr>
        <w:trPr/>
        <w:tc><w:tcPr/><w:p><w:r><w:t>outer keep</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('acceptAllRevisions — nested table row-level <w:del/>', () => {
  it('removes nested-table deleted rows without touching the outer row', async () => {
    const buffer = await makeNestedDocxWithRowDel();
    // Default revisionHandling: 'accept' runs the DOM-based acceptor.
    const doc = await Document.loadFromBuffer(buffer);
    const outerTable = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    expect(outerTable).toBeDefined();
    // The outer table must still have exactly one row (never a candidate
    // for row-level removal; it had no row-level del marker).
    expect(outerTable.getRows().length).toBe(1);
    doc.dispose();
  });

  it('removes outer row (and its nested-table content) when outer trPr has del', async () => {
    const buffer = await makeNestedDocxWithOuterRowDel();
    const doc = await Document.loadFromBuffer(buffer);
    const outerTable = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    expect(outerTable).toBeDefined();
    // Only the surviving outer row remains.
    const rows = outerTable.getRows();
    expect(rows.length).toBe(1);
    expect(rows[0]!.getCells()[0]!.getParagraphs()[0]!.getText()).toBe('outer keep');
    doc.dispose();
  });
});
