/**
 * `stripTrackedChanges` must process every DOCX part that can carry
 * block-level content with tracked changes, not just word/document.xml.
 *
 * Per ECMA-376 Part 1, tracked-change markup (w:ins / w:del /
 * w:moveFrom / w:moveTo / *PrChange / *RangeStart / *RangeEnd) can
 * appear inside:
 *   - word/document.xml          (body content)
 *   - word/header*.xml           (§17.10.4 CT_HdrFtr)
 *   - word/footer*.xml           (§17.10.3 CT_HdrFtr)
 *   - word/footnotes.xml         (§17.11.15 w:footnote)
 *   - word/endnotes.xml          (§17.11.4 w:endnote)
 *   - word/comments.xml          (§17.13.4 w:comment)
 *
 * Bug: `stripTrackedChanges.ts` only walked `word/document.xml`.
 * Headers, footers, footnotes, endnotes, and comments kept their
 * tracked-change markup intact, so `stripTrackedChanges()` on a
 * document with footnote-level or header-level revisions left the
 * document in a partially-stripped state — Word still displayed the
 * remaining tracked changes.
 *
 * Iteration 136 extends the stripper to process every revision-
 * carrying part. Uses the same regex pipeline already used for
 * document.xml.
 */

import { ZipHandler } from '../../src/zip/ZipHandler';
import { stripTrackedChanges } from '../../src/utils/stripTrackedChanges';

function makeBaseZip() {
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
  zip.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body>
</w:document>`
  );
  return zip;
}

const revisionSnippet = `
      <w:ins w:id="10" w:author="A" w:date="2026-04-24T10:00:00Z">
        <w:r><w:t>ins</w:t></w:r>
      </w:ins>
      <w:del w:id="11" w:author="A" w:date="2026-04-24T10:00:00Z">
        <w:r><w:delText>del</w:delText></w:r>
      </w:del>`;

describe('stripTrackedChanges — non-body parts coverage', () => {
  it('strips revisions from word/header1.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/header1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>kept </w:t></w:r>${revisionSnippet}</w:p>
</w:hdr>`
    );
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/header1.xml') ?? '';
    expect(after).not.toMatch(/<w:ins\b|<w:del\b/);
    expect(after).toContain('kept');
    expect(after).toContain('ins');
    expect(after).not.toContain('<w:delText>del</w:delText>');
  });

  it('strips revisions from word/footer2.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/footer2.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>foot </w:t></w:r>${revisionSnippet}</w:p>
</w:ftr>`
    );
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/footer2.xml') ?? '';
    expect(after).not.toMatch(/<w:ins\b|<w:del\b/);
    expect(after).toContain('foot');
  });

  it('strips revisions from word/footnotes.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/footnotes.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>note </w:t></w:r>${revisionSnippet}</w:p>
  </w:footnote>
</w:footnotes>`
    );
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/footnotes.xml') ?? '';
    expect(after).not.toMatch(/<w:ins\b|<w:del\b/);
    expect(after).toContain('note');
  });

  it('strips revisions from word/endnotes.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/endnotes.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1">
    <w:p><w:r><w:t>end </w:t></w:r>${revisionSnippet}</w:p>
  </w:endnote>
</w:endnotes>`
    );
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/endnotes.xml') ?? '';
    expect(after).not.toMatch(/<w:ins\b|<w:del\b/);
    expect(after).toContain('end');
  });

  it('strips revisions from word/comments.xml', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/comments.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="1" w:author="A" w:date="2026-04-24T10:00:00Z">
    <w:p><w:r><w:t>c </w:t></w:r>${revisionSnippet}</w:p>
  </w:comment>
</w:comments>`
    );
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/comments.xml') ?? '';
    expect(after).not.toMatch(/<w:ins\b|<w:del\b/);
    expect(after).toContain('c');
  });

  it('strips pPrChange from a header without destroying unrelated content (regression guard)', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/header1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:pPrChange w:id="20" w:author="A" w:date="2026-04-24T10:00:00Z">
        <w:pPr><w:jc w:val="left"/></w:pPr>
      </w:pPrChange>
    </w:pPr>
    <w:r><w:t>centered</w:t></w:r>
  </w:p>
</w:hdr>`
    );
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/header1.xml') ?? '';
    expect(after).not.toMatch(/<w:pPrChange\b/);
    expect(after).toContain('centered');
    // Current alignment preserved
    expect(after).toContain('w:val="center"');
  });

  it('leaves a header without revisions unchanged (regression guard)', async () => {
    const zip = makeBaseZip();
    zip.addFile(
      'word/header1.xml',
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>no-revisions</w:t></w:r></w:p>
</w:hdr>`
    );
    await stripTrackedChanges(zip);
    const after = zip.getFileAsString('word/header1.xml') ?? '';
    expect(after).toContain('no-revisions');
    expect(after).not.toMatch(/<w:ins\b|<w:del\b/);
  });
});
