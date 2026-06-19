/**
 * getHyperlinks()/getBookmarks()/getFields() must report table-cell content
 * exactly once.
 *
 * getAllParagraphs() already walks table-cell paragraphs (via walkElements),
 * so the previous extra getTables() pass pushed every table hyperlink/bookmark/
 * field a second time. updateAllHyperlinks() consumes getHyperlinks(), so a
 * non-idempotent formatter ran twice per table hyperlink and the returned count
 * was inflated. Removing the redundant table loops fixes the duplication.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function createDocxWithTableContent(): Promise<Buffer> {
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
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/cell" TargetMode="External"/>
</Relationships>`
  );

  // One hyperlink, one bookmark, and one MERGEFIELD — all inside a single table cell.
  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
      <w:tr>
        <w:tc>
          <w:p>
            <w:bookmarkStart w:id="1" w:name="CellMark"/>
            <w:hyperlink r:id="rId5"><w:r><w:t>Cell Link</w:t></w:r></w:hyperlink>
            <w:bookmarkEnd w:id="1"/>
          </w:p>
          <w:p>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText xml:space="preserve"> MERGEFIELD Name </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>Name</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

describe('Table content is not double-counted by collection accessors', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('getHyperlinks() reports a table-cell hyperlink exactly once', async () => {
    const buffer = await createDocxWithTableContent();
    doc = await Document.loadFromBuffer(buffer);

    expect(doc.getHyperlinks().length).toBe(1);
  });

  it('getBookmarks() reports a table-cell bookmark exactly once', async () => {
    const buffer = await createDocxWithTableContent();
    doc = await Document.loadFromBuffer(buffer);

    expect(doc.getBookmarks().length).toBe(1);
  });

  it('getFields() reports a table-cell field exactly once', async () => {
    const buffer = await createDocxWithTableContent();
    doc = await Document.loadFromBuffer(buffer);

    expect(doc.getFields().length).toBe(1);
  });

  it('updateAllHyperlinks() invokes the formatter once with the correct table hyperlink', async () => {
    const buffer = await createDocxWithTableContent();
    doc = await Document.loadFromBuffer(buffer);

    const seenUrls: (string | undefined)[] = [];
    const count = doc.updateAllHyperlinks((link) => {
      seenUrls.push(link.getUrl());
      link.setFormatting({ color: 'FF0000', bold: true });
    });

    // Exactly one invocation, with the actual cell hyperlink (not a duplicate
    // or the wrong object), and the returned count reflects that.
    expect(seenUrls).toEqual(['https://example.com/cell']);
    expect(count).toBe(1);

    // The mutation applied through the callback is observable on the model and
    // is not lost (the formatter operated on the live hyperlink instance).
    const runFmt = doc.getHyperlinks()[0]!.hyperlink.getRun().getFormatting();
    expect(runFmt.color).toBe('FF0000');
    expect(runFmt.bold).toBe(true);
  });
});
