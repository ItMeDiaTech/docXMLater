/**
 * Raw-XML revision acceptor — row-level `<w:del/>` must remove the row.
 *
 * Per ECMA-376 Part 1 §17.13.5.14, `<w:del/>` inside `<w:trPr>` marks the
 * entire row as a tracked deletion. Accepting the deletion must remove the
 * `<w:tr>...</w:tr>` block, not just the marker.
 *
 * Bug guarded against: `acceptRevisions.ts#acceptDeletions` strips the
 * self-closing `<w:del/>` marker but leaves the `<w:tr>` wrapper and its
 * (now-empty) cells in place. Documents loaded with `revisionHandling:
 * 'accept'` (the default) accumulate zombie rows with no visible content,
 * violating the semantic contract of tracked-change acceptance.
 */

import { Document } from '../../src/core/Document';
import { Table } from '../../src/elements/Table';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithDeletedRow(): Promise<Buffer> {
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
          <w:del w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z"/>
        </w:trPr>
        <w:tc><w:tcPr/><w:p><w:r><w:t>gone</w:t></w:r></w:p></w:tc>
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

describe('acceptAllRevisions (raw-XML path) — row-level <w:del/>', () => {
  it('removes the entire <w:tr> when its <w:trPr> contains <w:del/>', async () => {
    const buffer = await makeDocxWithDeletedRow();
    // Default revisionHandling: 'accept' invokes the raw-XML acceptor.
    const doc = await Document.loadFromBuffer(buffer);
    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    expect(table).toBeDefined();
    const rows = table.getRows();
    // The deleted row must be removed; only 'keep' and 'tail' remain.
    expect(rows.length).toBe(2);
    const texts = rows.map((r) => r.getCells()[0]!.getParagraphs()[0]!.getText());
    expect(texts).toEqual(['keep', 'tail']);
    doc.dispose();
  });

  it('does not remove rows containing <w:del/> elsewhere (e.g. wrapping run content inside a cell)', async () => {
    // Build a doc where <w:del> wraps a run in a cell — NOT a row-level marker.
    // The row must survive.
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
        <w:tc>
          <w:tcPr/>
          <w:p>
            <w:del w:id="1" w:author="A" w:date="2026-01-15T10:00:00Z">
              <w:r><w:delText>deleted text</w:delText></w:r>
            </w:del>
            <w:r><w:t> surviving</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p/>
  </w:body>
</w:document>`
    );
    const buffer = await zipHandler.toBuffer();
    const doc = await Document.loadFromBuffer(buffer);
    const table = doc.getBodyElements().find((el) => el instanceof Table) as Table;
    const rows = table.getRows();
    // Row must survive because the <w:del> was wrapping content, not row-level.
    expect(rows.length).toBe(1);
    const text = rows[0]!.getCells()[0]!.getParagraphs()[0]!.getText();
    // Deleted content stripped; surviving content kept.
    expect(text).toBe(' surviving');
    doc.dispose();
  });
});
