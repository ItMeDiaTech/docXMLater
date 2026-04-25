/**
 * `acceptAllRevisions` (raw-XML) must process footnotes.xml,
 * endnotes.xml, and comments.xml in addition to document.xml and
 * headers/footers.
 *
 * Per ECMA-376 Part 1:
 *   - Footnotes (§17.11.15 w:footnote) can contain any block-level
 *     content including tracked changes (w:ins / w:del / w:moveFrom /
 *     w:moveTo / pPrChange / rPrChange).
 *   - Endnotes (§17.11.4 w:endnote) — same content model.
 *   - Comments (§17.13.4 w:comment) — same content model.
 *
 * Bug: `acceptRevisions.ts` only walked `word/document.xml`,
 * `word/header*.xml`, and `word/footer*.xml`. Tracked changes inside
 * footnotes, endnotes, and comments were silently left in place,
 * meaning `acceptAllRevisions()` on a document with footnote-level
 * revisions returned without fully accepting all of the revisions —
 * the resulting document still had unaccepted tracked changes visible
 * in Word.
 *
 * Iteration 135 extends the raw-XML acceptor to also process
 * footnotes.xml, endnotes.xml, and comments.xml.
 */

import { ZipHandler } from '../../src/zip/ZipHandler';
import { acceptAllRevisions } from '../../src/utils/acceptRevisions';

function makeBaseZip() {
  const zip = new ZipHandler();
  zip.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`
  );
  zip.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );
  zip.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return zip;
}

describe('acceptAllRevisions(raw-XML) — footnote / endnote / comment coverage', () => {
  it('accepts a tracked insertion inside word/footnotes.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/footnotes.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p>
      <w:r><w:t>kept </w:t></w:r>
      <w:ins w:id="100" w:author="A" w:date="2026-04-24T10:00:00Z">
        <w:r><w:t>inserted</w:t></w:r>
      </w:ins>
    </w:p>
  </w:footnote>
</w:footnotes>`
    );
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/footnotes.xml') ?? '';
    // Wrapper gone, content preserved.
    expect(after).not.toMatch(/<w:ins\b/);
    expect(after).toContain('inserted');
    expect(after).toContain('kept');
  });

  it('accepts a tracked deletion inside word/footnotes.xml (content removed)', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/footnotes.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p>
      <w:r><w:t>kept </w:t></w:r>
      <w:del w:id="101" w:author="A" w:date="2026-04-24T10:00:00Z">
        <w:r><w:delText>removed</w:delText></w:r>
      </w:del>
    </w:p>
  </w:footnote>
</w:footnotes>`
    );
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/footnotes.xml') ?? '';
    expect(after).not.toMatch(/<w:del\b/);
    expect(after).not.toContain('removed');
    expect(after).toContain('kept');
  });

  it('accepts a tracked insertion inside word/endnotes.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/endnotes.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1">
    <w:p>
      <w:ins w:id="200" w:author="A" w:date="2026-04-24T10:00:00Z">
        <w:r><w:t>end-ins</w:t></w:r>
      </w:ins>
    </w:p>
  </w:endnote>
</w:endnotes>`
    );
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/endnotes.xml') ?? '';
    expect(after).not.toMatch(/<w:ins\b/);
    expect(after).toContain('end-ins');
  });

  it('accepts a tracked deletion inside word/comments.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/comments.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="A" w:date="2026-04-24T10:00:00Z">
    <w:p>
      <w:r><w:t>kept </w:t></w:r>
      <w:del w:id="300" w:author="B" w:date="2026-04-24T11:00:00Z">
        <w:r><w:delText>removed-from-comment</w:delText></w:r>
      </w:del>
    </w:p>
  </w:comment>
</w:comments>`
    );
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/comments.xml') ?? '';
    expect(after).not.toMatch(/<w:del\b/);
    expect(after).not.toContain('removed-from-comment');
    expect(after).toContain('kept');
  });

  it('leaves footnotes.xml untouched when there are no revisions (regression guard)', async () => {
    const zip = makeBaseZip();
    const clean = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:t>pure</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;
    zip.addFile('word/footnotes.xml', clean);
    await acceptAllRevisions(zip);
    const after = zip.getFileAsString('word/footnotes.xml') ?? '';
    expect(after).toContain('pure');
    expect(after).not.toMatch(/<w:ins\b|<w:del\b/);
  });
});
