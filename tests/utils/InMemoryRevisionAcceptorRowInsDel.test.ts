/**
 * InMemoryRevisionAcceptor — row-level tracked ins / del acceptance.
 *
 * Per ECMA-376 Part 1 §17.13.5.14 / §17.13.5.19:
 *   - `<w:trPr><w:ins/></w:trPr>` marks the entire row as a tracked insertion.
 *     Accepting insertions keeps the row and clears the marker.
 *   - `<w:trPr><w:del/></w:trPr>` marks the entire row as a tracked deletion.
 *     Accepting deletions REMOVES the row from the table.
 *
 * Bug guarded against: `acceptRevisionsInMemory` walked tables to clear
 * `trPrChange`, `tcPrChange`, and `cellRevision` but ignored the
 * `rowInsertion` / `rowDeletion` fields added in the previous iteration.
 * Consequently a user could load a doc with `revisionHandling: 'preserve'`,
 * call `acceptAllRevisions()`, and find deleted rows still present and
 * inserted rows still marked as pending — tracked changes were silently
 * not accepted on the row level.
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { acceptRevisionsInMemory } from '../../src/processors/InMemoryRevisionAcceptor';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithTrackedRows(): Promise<Buffer> {
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
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
      <w:tr>
        <w:trPr/>
        <w:tc><w:tcPr/><w:p><w:r><w:t>keep</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:trPr>
          <w:ins w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z"/>
        </w:trPr>
        <w:tc><w:tcPr/><w:p><w:r><w:t>inserted</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:trPr>
          <w:del w:id="2" w:author="A" w:date="2026-01-15T10:00:00Z"/>
        </w:trPr>
        <w:tc><w:tcPr/><w:p><w:r><w:t>deleted</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:trPr/>
        <w:tc><w:tcPr/><w:p><w:r><w:t>tail</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
  </w:body>
</w:document>`
  );
  return await zipHandler.toBuffer();
}

describe('acceptAllRevisions — row-level w:ins / w:del', () => {
  it('clears rowInsertion markers when accepting insertions, keeps the row', async () => {
    const buffer = await makeDocxWithTrackedRows();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });
    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;

    // Pre-condition: inserted row (index 1) has rowInsertion set.
    expect(table.getRows()[1]!.getFormatting().rowInsertion).toBeDefined();

    acceptRevisionsInMemory(doc, {
      acceptInsertions: true,
      acceptDeletions: false,
      acceptMoves: false,
      acceptPropertyChanges: false,
    });

    const reTable = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const rows = reTable.getRows();
    // Inserted row is still present but its marker is cleared.
    const insertedRow = rows[1]!;
    expect(insertedRow.getFormatting().rowInsertion).toBeUndefined();
    // Deleted row still present (we didn't accept deletions yet).
    expect(rows.length).toBe(4);
    doc.dispose();
  });

  it('removes rows with rowDeletion when accepting deletions', async () => {
    const buffer = await makeDocxWithTrackedRows();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    acceptRevisionsInMemory(doc, {
      acceptInsertions: false,
      acceptDeletions: true,
      acceptMoves: false,
      acceptPropertyChanges: false,
    });

    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const rows = table.getRows();
    // Deleted row removed; only 3 rows remain (keep, inserted, tail).
    expect(rows.length).toBe(3);
    // Remaining rows should not include the deleted-row's content.
    const texts = rows.map((r) => r.getCells()[0]!.getParagraphs()[0]!.getText());
    expect(texts).toEqual(['keep', 'inserted', 'tail']);
    doc.dispose();
  });

  it('accepts both insertions and deletions in one pass', async () => {
    const buffer = await makeDocxWithTrackedRows();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    acceptRevisionsInMemory(doc, {
      acceptInsertions: true,
      acceptDeletions: true,
      acceptMoves: false,
      acceptPropertyChanges: false,
    });

    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const rows = table.getRows();
    expect(rows.length).toBe(3);
    // Inserted row's marker cleared; deleted row removed.
    expect(rows[1]!.getFormatting().rowInsertion).toBeUndefined();
    expect(rows[1]!.getFormatting().rowDeletion).toBeUndefined();
    doc.dispose();
  });

  it('leaves markers intact when neither flag is set (acceptPropertyChanges only)', async () => {
    const buffer = await makeDocxWithTrackedRows();
    const doc = await Document.loadFromBuffer(buffer, { revisionHandling: 'preserve' });

    acceptRevisionsInMemory(doc, {
      acceptInsertions: false,
      acceptDeletions: false,
      acceptMoves: false,
      acceptPropertyChanges: true,
    });

    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const rows = table.getRows();
    expect(rows.length).toBe(4);
    expect(rows[1]!.getFormatting().rowInsertion).toBeDefined();
    expect(rows[2]!.getFormatting().rowDeletion).toBeDefined();
    doc.dispose();
  });
});
